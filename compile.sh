#!/bin/sh
# Compile SiPe.Sat into a Linux executable
pyinstaller main.py --clean --log-level INFO --onefile --windowed --name sipesat-final-second-draft-20230726
