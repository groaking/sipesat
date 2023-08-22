#!/bin/sh
# Compile SiPe.Sat into a Linux executable
pyinstaller main.py --clean --log-level INFO --onefile --windowed --name sipesat-final-third-draft-20230822
