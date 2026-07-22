import os
import sys


def pre_find_module_path(hook_api):
    """Use the installed tkinter package when automatic Tcl probing fails."""
    hook_api.search_dirs = [os.path.join(sys.base_prefix, "Lib")]
