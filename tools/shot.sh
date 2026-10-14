#!/bin/bash
# usage: tools/shot.sh <app> <out.png> [eval.js] [extra args...]
cd "$(dirname "$0")/.."
APP=$1; OUT=$2; EV=$3; shift 3 2>/dev/null
BTS_SHOT="$OUT" BTS_EVAL="$EV" xvfb-run -a -s "-screen 0 1600x1000x24" ./node_modules/.bin/electron . --no-sandbox --disable-gpu --app=$APP --user-data-dir=/tmp/claude-0/ud-$APP-$$ "$@" 2>&1 | grep -v -E "dbus|Fontconfig|libva|GPU|gles|viz|ERROR:bus|gbm|Exiting GPU|sandbox" | head -60
