#!/bin/sh
# PUID / PGID が指定されたら、Apache の実行ユーザー www-data の UID/GID を
# 起動時に合わせる（フォルダ所有者と揃えることで権限エラーを回避する）。
# 未指定なら従来通り www-data (33:33) で動く。
# usermod が無い・失敗する環境でも効くよう、/etc/passwd・/etc/group の
# 直接書き換えをフォールバックにする。結果はコンテナログに出す。
set -e
echo "[ems-entrypoint] PUID=${PUID:-} PGID=${PGID:-}" >&2
if [ -n "${PGID:-}" ] && [ "$PGID" != "33" ]; then
    if command -v groupmod >/dev/null 2>&1; then
        groupmod -g "$PGID" www-data 2>/dev/null || true
    fi
    if [ "$(id -g www-data 2>/dev/null)" != "$PGID" ]; then
        sed -i "s/^www-data:x:[0-9]*:/www-data:x:${PGID}:/" /etc/group 2>/dev/null || true
    fi
fi
if [ -n "${PUID:-}" ] && [ "$PUID" != "33" ]; then
    if command -v usermod >/dev/null 2>&1; then
        usermod -u "$PUID" www-data 2>/dev/null || true
    fi
    if [ "$(id -u www-data 2>/dev/null)" != "$PUID" ]; then
        CURG=$(id -g www-data 2>/dev/null || echo 33)
        sed -i "s/^www-data:x:[0-9]*:[0-9]*:/www-data:x:${PUID}:${CURG}:/" /etc/passwd 2>/dev/null || true
    fi
fi
echo "[ems-entrypoint] www-data: $(id www-data 2>/dev/null || echo unknown)" >&2
# 索引フォルダは実行ユーザーで書けるようにする
mkdir -p /Index 2>/dev/null || true
chown www-data:www-data /Index 2>/dev/null || true
exec "$@"
