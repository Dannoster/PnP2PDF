#!/bin/sh
pyinstaller \
    --add-data "pnp2pdf/color_profiles/sRGB-IEC61966-2.1.icc:color_profiles" \
    --add-data "pnp2pdf/color_profiles/USWebCoatedSWOP.icc:color_profiles" \
    --add-data "pnp2pdf/icon.png:." \
    --icon=pnp2pdf/icon.png \
    --windowed PnP2PDF.py
