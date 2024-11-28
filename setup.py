# import sys
# import os
# from cx_Freeze import setup, Executable
#
# # Determine base according to the platform
# if sys.platform == "win32":
#     # For Windows, use GUI mode without a console window
#     base = "Win32GUI"
# elif sys.platform == "darwin":  # macOS
#     # macOS doesn't require a specific base for GUI
#     base = None
# else:
#     # Default base for other platforms
#     base = None
#
# # Absolute path for the model cache directory
# model_cache_path = os.path.abspath("model_cache/")
#
# # Build options for cx_Freeze
# build_exe_options = {
#     # List of packages to include in the build
#     "packages": [
#         "tkinter", "transformers", "torch", "torchvision",
#         "matplotlib", "pptx", "os", "csv", "re", "datetime", "pathlib"
#     ],
#     # Files or directories to include in the build
#     "include_files": [
#         model_cache_path,  # Include the model cache directory
#     ],
#     # Modules to exclude (if any)
#     "excludes": [],
#     # Optional: Specify the output directory for build files
#     "build_exe": "dist",
# }
#
# # Define the executable
# executables = [
#     Executable(
#         "main.py",  # Entry point script for the application
#         base=base,  # Base type (GUI or console)
#         target_name="ppt_validator.exe" if sys.platform == "win32" else "ppt_validator"  # Name of the output executable
#     )
# ]
#
# # Setup configuration
# setup(
#     name="PPT Validator",  # Application name
#     version="1.0",  # Version number
#     description="PowerPoint Grammar and Font Validation Tool",  # Short description of the application
#     options={"build_exe": build_exe_options},  # Build options
#     executables=executables,  # Executables to generate
# )

import sys
import os
from cx_Freeze import setup, Executable

# Tentukan base tergantung pada platform
if sys.platform == "win32":
    base = "Win32GUI"
else:
    base = None

# Path absolut untuk direktori cache model
model_cache_path = os.path.abspath("model_cache/")
additional_includes = [
    ("model_cache", "model_cache"),  # Tambahkan cache model
]

# Opsi build untuk cx_Freeze
build_exe_options = {
    "packages": [
        "tkinter", "transformers", "torch", "torchvision", "matplotlib", "pptx",
        "os", "csv", "re", "datetime", "pathlib", "logging"
    ],
    "include_files": additional_includes,
    "excludes": [],
    "build_exe": "dist",  # Folder keluaran
}

# Tentukan executable
executables = [
    Executable(
        "main.py",
        base=base,
        target_name="ppt_validator.exe"
    )
]

# Konfigurasi setup
setup(
    name="PPT Validator",
    version="1.0",
    description="PowerPoint Grammar and Font Validation Tool",
    options={"build_exe": build_exe_options},
    executables=executables,
)
