#!/bin/sh
# PUID / PGID が指定されたら、Apache の実行ユーザー www-data の UID/GID を
# 起動時に合わせる（フォルダ所有者と揃えることで権限エラーを回避する）。
# 未指定なら従来通り www-data (33:33) で動く。
set -e
if [ -n "${PGID:-}" ] && [ "$PGID" != "33" ]; then
    groupmod -g "$PGID" www-data 2>/dev/null || true
fi
if [ -n "${PUID:-}" ] && [ "$PUID" != "33" ]; then
    usermod -u "$PUID" www-data 2>/dev/null || true
fi
# 索引フォルダは実行ユーザーで書けるようにする
mkdir -p /Index 2>/dev/null || true
chown www-data:www-data /Index 2>/dev/null || true
exec "$@"
