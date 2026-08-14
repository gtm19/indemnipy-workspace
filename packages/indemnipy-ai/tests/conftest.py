import warnings

# oletools calls deprecated pyparsing APIs at import time, before pytest's ini
# filterwarnings take effect.
warnings.filterwarnings("ignore", category=DeprecationWarning, module=r"oletools")
