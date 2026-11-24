#!/bin/sh
# PUID / PGID が指定されたら、Apache の実行ユーザー www-data の UID/GID を
# 起動時に合わせる（フォルダ所有者と揃えることで権限エラーを回避する）。
# 未指定なら従来通り www-data (33:33) で動く。
# /etc/passwd と /etc/group の www-data 行を直接書き換える方式
# （usermod だと passwd 側の GID が古いまま残るケースがあったため）。
# 結果はコンテナログに出す。
set -e
echo "[ems-entrypoint] PUID=${PUID:-} PGID=${PGID:-}" >&2
if [ -n "${PUID:-}" ] || [ -n "${PGID:-}" ]; then
    NEWUID=${PUID:-33}
    NEWGID=${PGID:-33}
    sed -i "s/^www-data:x:[0-9]*:[0-9]*:/www-data:x:${NEWUID}:${NEWGID}:/" /etc/passwd 2>/dev/null || true
    sed -i "s/^www-data:x:[0-9]*:/www-data:x:${NEWGID}:/" /etc/group 2>/dev/null || true
fi
echo "[ems-entrypoint] www-data: $(id www-data 2>/dev/null || echo unknown)" >&2
# 索引フォルダは実行ユーザーで書けるようにする
mkdir -p /Index 2>/dev/null || true
chown www-data:www-data /Index 2>/dev/null || true
exec "$@"
