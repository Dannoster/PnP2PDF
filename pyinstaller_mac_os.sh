#!/bin/sh
# Kept for compatibility: the full build lives in build_app.sh.
exec "$(dirname "$0")/build_app.sh" "$@"
