# Excel_Multi-Sheet_Search
# ビルド: docker compose build
# 使い方: docker-compose.yml の volumes で Excelフォルダを /Work、
#         索引用フォルダを /Index に渡すこと
FROM php:8.3-apache

# zip 拡張（無くても内蔵リーダーで動くが、ある方が高速）
RUN apt-get update \
 && apt-get install -y --no-install-recommends libzip-dev \
 && docker-php-ext-install zip \
 && rm -rf /var/lib/apt/lists/*

# 日本語ファイル名・検索のためのロケール（ファイル名はバイト処理でも動くが念のため）
RUN apt-get update \
 && apt-get install -y --no-install-recommends locales \
 && echo 'ja_JP.UTF-8 UTF-8' >> /etc/locale.gen \
 && locale-gen ja_JP.UTF-8 \
 && rm -rf /var/lib/apt/lists/*
ENV LANG=ja_JP.UTF-8 LC_ALL=ja_JP.UTF-8

COPY index.php /var/www/html/index.php
# ビルド元のファイル権限に引きずられないよう、Webから読める権限に固定する
RUN chmod 644 /var/www/html/index.php
COPY docker-entrypoint.sh /usr/local/bin/docker-entrypoint.sh
RUN chmod +x /usr/local/bin/docker-entrypoint.sh
ENTRYPOINT ["/usr/local/bin/docker-entrypoint.sh"]
CMD ["apache2-foreground"]

# 既定の受け渡し先（compose 側の volumes で上書きマウントされる）
RUN mkdir -p /Work /Index && chown -R www-data:www-data /Index
ENV EMS_WORK_DIR=/Work EMS_INDEX_DIR=/Index

EXPOSE 80
