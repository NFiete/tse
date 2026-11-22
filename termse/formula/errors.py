"""Excel error values."""
from __future__ import annotations


class ExcelError:
    """An Excel error value such as #DIV/0!. Instances are singletons per code."""

    __slots__ = ("code",)
    _registry: dict[str, "ExcelError"] = {}

    def __new__(cls, code: str):
        inst = cls._registry.get(code)
        if inst is None:
            inst = object.__new__(cls)
            inst.code = code
            cls._registry[code] = inst
        return inst

    def __repr__(self) -> str:
        return self.code

    __str__ = __repr__

    def __reduce__(self):
        return (ExcelError, (self.code,))


NULL = ExcelError("#NULL!")
DIV0 = ExcelError("#DIV/0!")
VALUE = ExcelError("#VALUE!")
REF = ExcelError("#REF!")
NAME = ExcelError("#NAME?")
NUM = ExcelError("#NUM!")
NA = ExcelError("#N/A")
SPILL = ExcelError("#SPILL!")
CALC = ExcelError("#CALC!")
GETTING_DATA = ExcelError("#GETTING_DATA")
# Not a real Excel error: shown for cells that are part of a circular reference.
CIRC = ExcelError("#CIRC!")

ERROR_CODES = {e.code: e for e in (NULL, DIV0, VALUE, REF, NAME, NUM, NA, SPILL, CALC, GETTING_DATA)}

# ERROR.TYPE numbering
ERROR_TYPE_NUM = {NULL: 1, DIV0: 2, VALUE: 3, REF: 4, NAME: 5, NUM: 6, NA: 7, GETTING_DATA: 8}


class XLError(Exception):
    """Raised inside the evaluator to short-circuit with an Excel error value."""

    __slots__ = ("err",)

    def __init__(self, err: ExcelError):
        self.err = err


class Unsupported(Exception):
    """Raised for syntax or functions this engine does not implement."""
