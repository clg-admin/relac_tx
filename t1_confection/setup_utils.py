# support_main_funs.py

# This file holds multiple functions to support the main operation of the "simulatod_cba_dlc" project.

import subprocess

def install_requirements(requirements_file):
    with open(requirements_file, 'r') as file:
        requirements = file.readlines()
    requirements = [req.strip() for req in requirements if req.strip()]

    for requirement in requirements:
        subprocess.call(['pip', 'install', requirement])

