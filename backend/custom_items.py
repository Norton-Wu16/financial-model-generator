"""自定义行项表达式解析器 - 手写词法 + 语法分析器。

不使用 eval/exec，通过白名单函数确保安全。

支持的语法:
  - 数字: 100, 3.14, -5
  - 运算符: + - * / ^ ( ^ 为幂运算，对应 Excel 的 ^)
  - 括号: ( )
  - 逗号: ,
  - 标识符: 函数名或变量名 (字母数字下划线，必须以字母或下划线开头)
  - 函数白名单: MAX, MIN, ABS, ROUND, SUM (大小写不敏感)
  - 变量名: 任意标识符 (在 context 中查找，找不到则报错)

示例:
  evaluate("revenue * 0.15", {"revenue": 100})  -> 15.0
  evaluate("MAX(net_income, 0)", {"net_income": -5})  -> 0.0
  evaluate("ebitda - ebit", {"ebitda": 30, "ebit": 25})  -> 5.0

安全策略:
  - 拒绝非白名单函数 (如 __import__, eval, exec, os.system)
  - 不支持属性访问 (.)、字符串字面量、方括号、赋值
  - 词法分析器遇到非法字符直接报错
"""
from typing import Any, Dict, List, Tuple, Union
from pydantic import BaseModel, Field


# ==================== Pydantic Schema ====================

class CustomItem(BaseModel):
    """自定义行项定义。"""
    name: str = Field(..., description="行项名称，如 '研发费用率'")
    formula: str = Field(..., description="计算公式，如 'rd / revenue * 100' 或 'MAX(net_income, 0)'")
    target_sheet: str = Field("", description="目标 sheet 名（可选，留空则使用模型默认 sheet）")
    format: str = Field("number", description="显示格式: 'number' 或 'percent'")


# ==================== 词法分析 Token ====================

# Token 类型
T_NUMBER = "NUMBER"
T_OP = "OP"           # + - * / ^
T_LPAREN = "LPAREN"
T_RPAREN = "RPAREN"
T_COMMA = "COMMA"
T_IDENT = "IDENT"     # 函数名或变量名
T_EOF = "EOF"


class Token:
    __slots__ = ("type", "value", "pos")

    def __init__(self, type_: str, value: Any, pos: int):
        self.type = type_
        self.value = value
        self.pos = pos

    def __repr__(self):
        return f"Token({self.type}, {self.value!r}, pos={self.pos})"


# 白名单函数（大写形式）
WHITELIST_FUNCS = {"MAX", "MIN", "ABS", "ROUND", "SUM"}


class LexError(Exception):
    """词法分析错误。"""
    pass


class ParseError(Exception):
    """语法分析错误。"""
    pass


class EvalError(Exception):
    """求值错误。"""
    pass


# ==================== 词法分析器 ====================

def tokenize(expr: str) -> List[Token]:
    """将表达式字符串切分为 token 列表。

    识别: 数字(整数/小数)、运算符(+ - * / ^)、括号、逗号、标识符
    不识别: 字符串、点号(.)、方括号、等号(=)、冒号(:) - 遇到即报错
    """
    tokens: List[Token] = []
    i = 0
    n = len(expr)
    while i < n:
        c = expr[i]

        # 跳过空白
        if c.isspace():
            i += 1
            continue

        # 数字: 整数或小数 (支持前导负号由运算符处理)
        if c.isdigit() or (c == "." and i + 1 < n and expr[i + 1].isdigit()):
            j = i
            dot_seen = False
            while j < n and (expr[j].isdigit() or expr[j] == "."):
                if expr[j] == ".":
                    if dot_seen:
                        break
                    dot_seen = True
                j += 1
            num_str = expr[i:j]
            try:
                val = float(num_str)
                # 整数则转 int 以保持整洁（不影响计算）
                if not dot_seen and num_str.isdigit():
                    val = int(num_str)
            except ValueError:
                raise LexError(f"无效数字: {num_str} (位置 {i})")
            tokens.append(Token(T_NUMBER, val, i))
            i = j
            continue

        # 标识符: 字母/下划线开头，后跟字母/数字/下划线
        if c.isalpha() or c == "_":
            j = i
            while j < n and (expr[j].isalnum() or expr[j] == "_"):
                j += 1
            name = expr[i:j]
            tokens.append(Token(T_IDENT, name, i))
            i = j
            continue

        # 运算符
        if c in "+-*/^":
            tokens.append(Token(T_OP, c, i))
            i += 1
            continue

        # 括号
        if c == "(":
            tokens.append(Token(T_LPAREN, c, i))
            i += 1
            continue
        if c == ")":
            tokens.append(Token(T_RPAREN, c, i))
            i += 1
            continue

        # 逗号
        if c == ",":
            tokens.append(Token(T_COMMA, c, i))
            i += 1
            continue

        # 显式拒绝危险字符（避免绕过）
        if c == ".":
            raise LexError(f"非法字符 '.' (属性访问不被支持)，位置 {i}")
        if c in "[]{}=:;\"'\\`~!@#$%&|<>?":
            raise LexError(f"非法字符 '{c}'，位置 {i}")

        raise LexError(f"非法字符 '{c}'，位置 {i}")

    tokens.append(Token(T_EOF, None, n))
    return tokens


# ==================== AST 节点 ====================
# 使用 dict 表示 AST 节点:
#   {"type": "num", "value": 100}
#   {"type": "var", "name": "revenue"}
#   {"type": "unary", "op": "-", "operand": <ast>}
#   {"type": "binop", "op": "+", "left": <ast>, "right": <ast>}
#   {"type": "call", "name": "MAX", "args": [<ast>, ...]}


# ==================== 递归下降语法分析器 ====================

class Parser:
    """递归下降语法分析器。

    文法:
      expr   := term (('+' | '-') term)*
      term   := factor (('*' | '/') factor)*
      factor := base ('^' factor)?      # 右结合
      base   := NUMBER
             | IDENT                       # 变量
             | IDENT '(' args ')'          # 函数调用
             | '(' expr ')'
             | ('+' | '-') base            # 一元正负号
      args   := expr (',' expr)*
    """

    def __init__(self, tokens: List[Token]):
        self.tokens = tokens
        self.pos = 0

    def peek(self) -> Token:
        return self.tokens[self.pos]

    def advance(self) -> Token:
        t = self.tokens[self.pos]
        self.pos += 1
        return t

    def expect(self, type_: str) -> Token:
        t = self.peek()
        if t.type != type_:
            raise ParseError(f"期望 {type_}，实际 {t.type} (位置 {t.pos})")
        return self.advance()

    def parse(self) -> Dict[str, Any]:
        node = self.parse_expr()
        if self.peek().type != T_EOF:
            t = self.peek()
            raise ParseError(f"意外的 token: {t.type} {t.value!r} (位置 {t.pos})")
        return node

    def parse_expr(self) -> Dict[str, Any]:
        left = self.parse_term()
        while self.peek().type == T_OP and self.peek().value in ("+", "-"):
            op = self.advance().value
            right = self.parse_term()
            left = {"type": "binop", "op": op, "left": left, "right": right}
        return left

    def parse_term(self) -> Dict[str, Any]:
        left = self.parse_factor()
        while self.peek().type == T_OP and self.peek().value in ("*", "/"):
            op = self.advance().value
            right = self.parse_factor()
            left = {"type": "binop", "op": op, "left": left, "right": right}
        return left

    def parse_factor(self) -> Dict[str, Any]:
        # ^ 右结合
        left = self.parse_base()
        if self.peek().type == T_OP and self.peek().value == "^":
            self.advance()
            right = self.parse_factor()  # 右结合: 递归调用 parse_factor
            left = {"type": "binop", "op": "^", "left": left, "right": right}
        return left

    def parse_base(self) -> Dict[str, Any]:
        t = self.peek()

        # 一元正负号
        if t.type == T_OP and t.value in ("+", "-"):
            op = self.advance().value
            operand = self.parse_base()
            return {"type": "unary", "op": op, "operand": operand}

        # 数字
        if t.type == T_NUMBER:
            self.advance()
            return {"type": "num", "value": t.value}

        # 括号
        if t.type == T_LPAREN:
            self.advance()
            node = self.parse_expr()
            self.expect(T_RPAREN)
            return node

        # 标识符: 变量或函数调用
        if t.type == T_IDENT:
            self.advance()
            # 函数调用?
            if self.peek().type == T_LPAREN:
                self.advance()  # 消费 '('
                args = []
                if self.peek().type != T_RPAREN:
                    args.append(self.parse_expr())
                    while self.peek().type == T_COMMA:
                        self.advance()
                        args.append(self.parse_expr())
                self.expect(T_RPAREN)
                # 函数白名单检查（大小写不敏感）
                upper_name = t.value.upper()
                if upper_name not in WHITELIST_FUNCS:
                    raise ParseError(
                        f"函数 '{t.value}' 不在白名单中 (位置 {t.pos})。"
                        f"允许的函数: {', '.join(sorted(WHITELIST_FUNCS))}"
                    )
                return {"type": "call", "name": upper_name, "args": args, "orig_name": t.value}
            # 变量
            return {"type": "var", "name": t.value}

        raise ParseError(f"意外的 token: {t.type} {t.value!r} (位置 {t.pos})")


def parse(tokens: List[Token]) -> Dict[str, Any]:
    """语法分析入口。"""
    return Parser(tokens).parse()


# ==================== 求值器 ====================

def evaluate(ast: Dict[str, Any], context: Dict[str, float]) -> float:
    """对 AST 求值。

    context: 变量名 → 数值 的字典。变量找不到则抛 EvalError。
    """
    t = ast["type"]

    if t == "num":
        return float(ast["value"])

    if t == "var":
        name = ast["name"]
        if name not in context:
            raise EvalError(f"未知变量: '{name}'")
        val = context[name]
        if val is None:
            return 0.0
        return float(val)

    if t == "unary":
        v = evaluate(ast["operand"], context)
        return -v if ast["op"] == "-" else v

    if t == "binop":
        left = evaluate(ast["left"], context)
        right = evaluate(ast["right"], context)
        op = ast["op"]
        if op == "+":
            return left + right
        if op == "-":
            return left - right
        if op == "*":
            return left * right
        if op == "/":
            if right == 0:
                raise EvalError("除零错误")
            return left / right
        if op == "^":
            try:
                return float(left ** right)
            except (ValueError, OverflowError) as e:
                raise EvalError(f"幂运算错误: {e}")
        raise EvalError(f"未知运算符: {op}")

    if t == "call":
        name = ast["name"]
        args = [evaluate(a, context) for a in ast["args"]]
        if name == "MAX":
            if not args:
                raise EvalError("MAX() 至少需要 1 个参数")
            return float(max(args))
        if name == "MIN":
            if not args:
                raise EvalError("MIN() 至少需要 1 个参数")
            return float(min(args))
        if name == "ABS":
            if len(args) != 1:
                raise EvalError("ABS() 需要 1 个参数")
            return float(abs(args[0]))
        if name == "ROUND":
            if len(args) not in (1, 2):
                raise EvalError("ROUND() 需要 1 或 2 个参数")
            digits = int(args[1]) if len(args) == 2 else 0
            return float(round(args[0], digits))
        if name == "SUM":
            return float(sum(args))
        raise EvalError(f"未知函数: {name}")  # 理论上不会到达（白名单已过滤）

    raise EvalError(f"未知 AST 节点类型: {t}")


# ==================== Excel 公式生成器 ====================

def to_excel_formula(ast: Dict[str, Any], cell_map: Dict[str, str]) -> str:
    """将 AST 转换为 Excel 公式字符串（不含前导 =）。

    cell_map: 变量名 → Excel 单元格引用 (如 "revenue" -> "IS!B5" 或 "B5")
    """
    t = ast["type"]

    if t == "num":
        v = ast["value"]
        # 整数显示为整数
        if isinstance(v, float) and v.is_integer():
            return str(int(v))
        return str(v)

    if t == "var":
        name = ast["name"]
        if name not in cell_map:
            raise EvalError(f"变量 '{name}' 未在 cell_map 中找到单元格引用")
        return cell_map[name]

    if t == "unary":
        inner = to_excel_formula(ast["operand"], cell_map)
        op = ast["op"]
        return f"({op}{inner})"

    if t == "binop":
        left = to_excel_formula(ast["left"], cell_map)
        right = to_excel_formula(ast["right"], cell_map)
        op = ast["op"]
        # ^ 在 Excel 中就是幂运算符
        return f"({left}{op}{right})"

    if t == "call":
        name = ast["name"]
        args = [to_excel_formula(a, cell_map) for a in ast["args"]]
        return f"{name}({','.join(args)})"

    raise EvalError(f"未知 AST 节点类型: {t}")


# ==================== 高级 API ====================

def evaluate_expression(expr: str, context: Dict[str, float]) -> float:
    """一站式: tokenize -> parse -> evaluate。"""
    tokens = tokenize(expr)
    ast = parse(tokens)
    return evaluate(ast, context)


def expression_to_excel(expr: str, cell_map: Dict[str, str]) -> str:
    """一站式: tokenize -> parse -> to_excel_formula。返回含前导 = 的公式。"""
    tokens = tokenize(expr)
    ast = parse(tokens)
    return "=" + to_excel_formula(ast, cell_map)


def validate_expression(expr: str) -> Tuple[bool, str]:
    """校验表达式语法 + 安全性，返回 (是否合法, 错误信息或空串)。"""
    try:
        tokens = tokenize(expr)
        ast = parse(tokens)
        # 不 evaluate（无 context），仅校验语法 + 白名单
        return True, ""
    except (LexError, ParseError) as e:
        return False, str(e)


def get_referenced_vars(expr: str) -> List[str]:
    """返回表达式引用的所有变量名（去重，保持顺序）。"""
    tokens = tokenize(expr)
    ast = parse(tokens)
    vars_: List[str] = []
    seen = set()

    def _walk(node: Dict[str, Any]):
        t = node["type"]
        if t == "var":
            name = node["name"]
            if name not in seen:
                seen.add(name)
                vars_.append(name)
        elif t == "unary":
            _walk(node["operand"])
        elif t == "binop":
            _walk(node["left"])
            _walk(node["right"])
        elif t == "call":
            for a in node["args"]:
                _walk(a)

    _walk(ast)
    return vars_
