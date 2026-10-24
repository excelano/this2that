#!/bin/bash
# uninstall.sh - Remove the this2that commands from ~/bin.
#
# Author: David M. Anderson
# Built with AI assistance (Claude, Anthropic)
#
# Usage: ./uninstall.sh

set -euo pipefail

for name in md2docx docx2md csv2xlsx xlsx2csv sqlite2csv sqlite2xlsx sqlite2ods parquet2csv md2pptx; do
    target="$HOME/bin/$name"
    if [ -e "$target" ]; then
        rm -f "$target"
        echo "Removed: $target"
    else
        echo "Not installed: $target (nothing to remove)"
    fi
done

# The ~/bin PATH entry in .bashrc is shared by the other nursery tools and is
# left alone on purpose.
