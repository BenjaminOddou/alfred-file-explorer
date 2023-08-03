#!/bin/bash

# Install platform agnostic version of openpyxl (put the script inside the workflow folder)
mkdir -p lib
pip3 install openpyxl -t lib --upgrade