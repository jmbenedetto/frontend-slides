#!/usr/bin/env bash
# Compatibility shim for skills installers that still expect this file.
# This fork intentionally does not support bundled PDF export helpers.
set -euo pipefail

echo "export-pdf.sh is intentionally disabled in this fork."
echo "This skill is local-first and does not provide bundled PDF export automation."
exit 1
