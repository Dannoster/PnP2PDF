#!/bin/sh
pyinstaller \
    --add-data "color_profiles/sRGB-IEC61966-2.1.icc:color_profiles" \
    --add-data "color_profiles/USWebCoatedSWOP.icc:color_profiles" \
    --add-data "icon.png:." \
    --icon=icon.png \
    --windowed PnP2PDF.py