# shim after reorg
from gui.pages.cargas.carga_page import *
__all__ = [k for k in globals().keys() if not k.startswith("_")]
