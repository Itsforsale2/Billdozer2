import os
import importlib


VENDOR_PARSER_FOLDER = "vendor_parsers"


def discover_vendors():
    """
    Scans vendor_parsers/ for files like:  knife_river_parser.py
    Returns vendor names: ['Knife River']
    """
    vendors = []

    if not os.path.exists(VENDOR_PARSER_FOLDER):
        return vendors

    for filename in os.listdir(VENDOR_PARSER_FOLDER):
        if filename.endswith("_parser.py"):
            vendor_raw = filename.replace("_parser.py", "")
            vendor_name = vendor_raw.replace("_", " ").title()
            vendors.append(vendor_name)

    return vendors


def get_vendor_parser(vendor_name):
    """
    Dynamically imports and returns the correct parser function
    for the selected vendor.
    """

    # Convert "Knife River" → "knife_river"
    vendor_raw = vendor_name.lower().replace(" ", "_")
    module_name = f"{VENDOR_PARSER_FOLDER}.{vendor_raw}_parser"
    function_name = f"parse_{vendor_raw}_pdf"

    # Import vendor module
    try:
        module = importlib.import_module(module_name)
    except ImportError:
        raise RuntimeError(f"Cannot import vendor parser: {module_name}")

    # Locate parse function
    if not hasattr(module, function_name):
        raise RuntimeError(
            f"Parser module '{module_name}' has no function '{function_name}'"
        )

    return getattr(module, function_name)
