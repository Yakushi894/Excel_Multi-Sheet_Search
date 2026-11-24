<?php
/**
 * Excel_Multi-Sheet_Search（品番コード → プライス品番・タグ品番）
 * ============================================================
 * Docker での使い方（推奨）
 *   docker-compose.yml で Excelフォルダを /Work、索引用フォルダを /Index に
 *   渡してください（docker-compose.yml の例を同梱）。
 *     volumes:
 *       - /path/to/Excelフォルダ:/Work:ro
 *       - ./index:/Index
 *   ブラウザで開いて 品番コード（8桁、例: 12345678）を入れて検索します。
 *
 * 設定（画面下部の「設定」から変更でき、/Index/settings.json に保存されます）
 *   - 検索フォルダ：初期値は /Work（Dockerで渡したフォルダ）。日本語フォルダ可
 *   - 読み込むシート：何枚目か（初期値＝1枚目）／シート名の一部／全シート
 *     ※シート名はファイルごとに変わる前提で、既定は「1枚目」です
 *   - 検索列：検索対象の列を A〜Z で指定（例: N ／ A-Z ／ B,N,P）
 *     空欄＝見出し行から「メーカー品番(品番コード)」列を自動判定
 *     セル内に「12345678+12345679」のように複数品番が入っている場合も、
 *     数字の塊ごとに切り分けて照合するので、どちらでもヒットします
 *   - 表示列：結果に表示する列を A〜Z で指定（最大10個）
 *     空欄＝標準表示（プライス品番／品番コード／商品名／売価など）
 *   - 結合セル：結合範囲は左上の値を全体に展開してから検索・表示します
 *     （横結合の見出し、縦結合で複数行にまたがる値も拾います）
 *
 * 索引（インデックス）
 *   /Index/search_index_v2.json に索引を保持し、2回目以降は高速です。
 *   Excelの追加・更新・削除は変更分だけ自動で読み直します。
 *   シート設定を変えた場合は索引を作り直します（自動）。
 *
 * その他
 *   - 日本語フォルダ名・日本語ファイル名対応（UTF-8）
 *   - 外部ライブラリ不要（ZipArchive+DOM、zip拡張が無いPHPでは内蔵リーダー）
 *   - 旧形式 .xls は対象外（.xlsx で保存し直してください）
 *   - Dockerを使わない場合は、このファイルをExcelフォルダに置けば
 *     そのフォルダが検索対象になり、索引も同じ場所に作られます
 */

error_reporting(E_ALL & ~E_DEPRECATED);
ini_set('display_errors', '0');
if (function_exists('set_time_limit')) { @set_time_limit(180); }
if (function_exists('mb_internal_encoding')) { mb_internal_encoding('UTF-8'); }

$EXTS = array('xlsx', 'xlsm');

/* ---------------- 保存先（索引・設定）の決定 ---------------- */

function bootstrap_index_dir() {
    $env = getenv('EMS_INDEX_DIR');
    if ($env && is_dir($env) && is_writable($env)) return $env;
    if (is_dir('/Index') && is_writable('/Index')) return '/Index';
    if (is_writable(__DIR__)) return __DIR__;
    return sys_get_temp_dir();
}
$INDEX_DIR = bootstrap_index_dir();
$SETTINGS_FILE = $INDEX_DIR . '/settings.json';
$CACHE_FILE = $INDEX_DIR . '/search_index_v2.json';

/* ---------------- 設定 ---------------- */

function default_settings() {
    $base = getenv('EMS_WORK_DIR');
    if (!$base) $base = is_dir('/Work') ? '/Work' : __DIR__;
    return array(
        'base_dir'     => $base,
        'sheet_mode'   => 'index', // index=何枚目 / name=名前の一部 / all=全シート
        'sheet_index'  => 1,
        'sheet_name'   => '',
        'search_cols'  => '',      // 例 "N" "A-Z" "B,N,P"（空欄=見出し自動判定）
        'display_cols' => '',      // 例 "F,N,C,L,O"（最大10、空欄=標準表示）
    );
}
function load_settings($file) {
    $s = default_settings();
    if (is_file($file)) {
        $j = json_decode((string)@file_get_contents($file), true);
        if (is_array($j)) foreach ($s as $k => $v) if (isset($j[$k])) $s[$k] = $j[$k];
    }
    $s['sheet_index'] = max(1, (int)$s['sheet_index']);
    if (!in_array($s['sheet_mode'], array('index', 'name', 'all'), true)) $s['sheet_mode'] = 'index';
    return $s;
}
$SETTINGS = load_settings($SETTINGS_FILE);

/* 設定の保存・索引リセット（POST） */
if ($_SERVER['REQUEST_METHOD'] === 'POST' && isset($_POST['action'])) {
    if ($_POST['action'] === 'save_settings') {
        $SETTINGS['base_dir'] = trim((string)$_POST['base_dir']);
        $SETTINGS['sheet_mode'] = (string)$_POST['sheet_mode'];
        $SETTINGS['sheet_index'] = max(1, (int)$_POST['sheet_index']);
        $SETTINGS['sheet_name'] = trim((string)$_POST['sheet_name']);
        $SETTINGS['search_cols'] = trim((string)$_POST['search_cols']);
        $SETTINGS['display_cols'] = trim((string)$_POST['display_cols']);
        if (!in_array($SETTINGS['sheet_mode'], array('index', 'name', 'all'), true)) $SETTINGS['sheet_mode'] = 'index';
        @file_put_contents($SETTINGS_FILE, json_encode($SETTINGS, JSON_UNESCAPED_UNICODE | JSON_PRETTY_PRINT), LOCK_EX);
    } elseif ($_POST['action'] === 'rebuild') {
        @unlink($CACHE_FILE);
    }
    $loc = strtok($_SERVER['REQUEST_URI'], '?') . '?saved=1';
    header('Location: ' . $loc);
    exit;
}

$BASE_DIR = $SETTINGS['base_dir'];

/* ---------------- 小物関数 ---------------- */

function h($s) { return htmlspecialchars((string)$s, ENT_QUOTES, 'UTF-8'); }

/** 全角数字→半角（区切り文字は残す） */
function halfwidth_digits($s) {
    return strtr((string)$s, array(
        '０'=>'0','１'=>'1','２'=>'2','３'=>'3','４'=>'4',
        '５'=>'5','６'=>'6','７'=>'7','８'=>'8','９'=>'9',
    ));
}
/** 数字だけ取り出す */
function digits_only($s) { return preg_replace('/[^0-9]/', '', halfwidth_digits($s)); }

/** 品番コードの照合キー：8桁に左ゼロ埋め（Excel数値だと先頭0が落ちるため） */
function code_key($digits) {
    if ($digits === '') return '';
    if (strlen($digits) <= 8) return str_pad($digits, 8, '0', STR_PAD_LEFT);
    return $digits;
}

/** セル内の数字の塊を全部取り出す（"12345678+12345679" → ["12345678","12345679"]） */
function number_tokens($cell) {
    $t = halfwidth_digits($cell);
    if (!preg_match_all('/[0-9]+/', $t, $m)) return array();
    return $m[0];
}

/** 数値セルの表示整形：12345678.0 → "12345678" */
function fmt_num($v) {
    $v = trim((string)$v);
    if ($v === '') return '';
    if (is_numeric($v)) {
        $f = (float)$v;
        if ($f == floor($f) && abs($f) < 1e15) return (string)(int)$f;
        return rtrim(rtrim($v, '0'), '.');
    }
    return $v;
}
function fmt_money($v) {
    $v = trim((string)$v);
    if ($v === '' || !is_numeric($v)) return $v;
    return number_format((float)$v);
}
function col_to_num($col) {
    $n = 0;
    for ($i = 0; $i < strlen($col); $i++) $n = $n * 26 + (ord($col[$i]) - 64);
    return $n;
}
function num_to_col($n) {
    $s = '';
    while ($n > 0) { $m = ($n - 1) % 26; $s = chr(65 + $m) . $s; $n = (int)(($n - $m - 1) / 26); }
    return $s;
}

/**
 * 列指定文字列を列文字の配列にする。
 * 例: "N" / "A-Z" / "B,N,P" / "Ｆ，Ｎ"（全角・読点も可）
 * $max を超えた分は切り捨て、戻り値の2番目で知らせる。
 */
function parse_cols_spec($spec, $max) {
    $spec = trim((string)$spec);
    if ($spec === '') return array(array(), false);
    if (function_exists('mb_convert_kana')) $spec = mb_convert_kana($spec, 'a');
    $spec = strtoupper(str_replace(array('、', '・', ' ', '　', '/', ';'), ',', $spec));
    $cols = array(); $truncated = false;
    foreach (explode(',', $spec) as $tok) {
        $tok = trim($tok);
        if ($tok === '') continue;
        if (preg_match('/^([A-Z])-([A-Z])$/', $tok, $m)) {
            for ($n = col_to_num($m[1]); $n <= col_to_num($m[2]); $n++) $cols[] = num_to_col($n);
        } elseif (preg_match('/^[A-Z]{1,2}$/', $tok)) {
            $cols[] = $tok;
        }
    }
    $cols = array_values(array_unique($cols));
    if ($max > 0 && count($cols) > $max) { $cols = array_slice($cols, 0, $max); $truncated = true; }
    return array($cols, $truncated);
}

list($SEARCH_COLS, $searchTrunc) = parse_cols_spec($SETTINGS['search_cols'], 0);
list($DISPLAY_COLS, $displayTrunc) = parse_cols_spec($SETTINGS['display_cols'], 10);

/* ---------------- xlsx 解析（ZipArchive + DOM、zip拡張が無い環境では純PHPリーダー） ---------------- */

class MiniZip {
    private $data = null; private $entries = array();
    public function __construct($path) {
        $this->data = @file_get_contents($path);
        if ($this->data === false) return;
        $len = strlen($this->data);
        $eocd = -1;
        for ($i = $len - 22; $i >= max(0, $len - 65557); $i--) {
            if (substr($this->data, $i, 4) === "PK\x05\x06") { $eocd = $i; break; }
        }
        if ($eocd < 0) { $this->data = null; return; }
        $cdCount = unpack('v', substr($this->data, $eocd + 10, 2))[1];
        $pos = unpack('V', substr($this->data, $eocd + 16, 4))[1];
        for ($n = 0; $n < $cdCount; $n++) {
            if (substr($this->data, $pos, 4) !== "PK\x01\x02") break;
            $method = unpack('v', substr($this->data, $pos + 10, 2))[1];
            $compSize = unpack('V', substr($this->data, $pos + 20, 4))[1];
            $nameLen = unpack('v', substr($this->data, $pos + 28, 2))[1];
            $extraLen = unpack('v', substr($this->data, $pos + 30, 2))[1];
            $commentLen = unpack('v', substr($this->data, $pos + 32, 2))[1];
            $localOffset = unpack('V', substr($this->data, $pos + 42, 4))[1];
            $name = substr($this->data, $pos + 46, $nameLen);
            $this->entries[$name] = array($method, $compSize, $localOffset);
            $pos += 46 + $nameLen + $extraLen + $commentLen;
        }
    }
    public function ok() { return $this->data !== null; }
    public function getFromName($name) {
        if ($this->data === null || !isset($this->entries[$name])) return false;
        list($method, $compSize, $offset) = $this->entries[$name];
        if (substr($this->data, $offset, 4) !== "PK\x03\x04") return false;
        $nameLen = unpack('v', substr($this->data, $offset + 26, 2))[1];
        $extraLen = unpack('v', substr($this->data, $offset + 28, 2))[1];
        $raw = substr($this->data, $offset + 30 + $nameLen + $extraLen, $compSize);
        if ($method === 0) return $raw;
        if ($method === 8) { $out = @gzinflate($raw); return $out === false ? false : $out; }
        return false;
    }
}
class XlsxReader {
    private $zip = null; private $mini = null;
    public function __construct($path) {
        if (class_exists('ZipArchive')) {
            $z = new ZipArchive();
            if ($z->open($path) === true) { $this->zip = $z; return; }
        }
        $m = new MiniZip($path);
        if ($m->ok()) $this->mini = $m;
    }
    public function ok() { return $this->zip !== null || $this->mini !== null; }
    public function getFromName($name) {
        if ($this->zip !== null) return $this->zip->getFromName($name);
        if ($this->mini !== null) return $this->mini->getFromName($name);
        return false;
    }
    public function close() { if ($this->zip !== null) $this->zip->close(); }
}
function load_dom($xml) {
    if ($xml === false || $xml === null || $xml === '') return null;
    $doc = new DOMDocument();
    libxml_use_internal_errors(true);
    $ok = $doc->loadXML($xml);
    libxml_clear_errors();
    return $ok ? $doc : null;
}

/** 1ファイル分の対象行を取り出す。戻り値: array('rows'=>[...], 'error'=>null|string) */
function parse_xlsx($path, $settings) {
    $zip = new XlsxReader($path);
    if (!$zip->ok()) return array('rows' => array(), 'error' => 'ファイルを開けませんでした');

    $shared = array();
    $ssDoc = load_dom($zip->getFromName('xl/sharedStrings.xml'));
    if ($ssDoc) foreach ($ssDoc->getElementsByTagName('si') as $si) $shared[] = $si->textContent;

    $wbDoc = load_dom($zip->getFromName('xl/workbook.xml'));
    if (!$wbDoc) { $zip->close(); return array('rows' => array(), 'error' => 'workbook.xml が読めません'); }
    $rels = array();
    $relDoc = load_dom($zip->getFromName('xl/_rels/workbook.xml.rels'));
    if ($relDoc) foreach ($relDoc->getElementsByTagName('Relationship') as $r) {
        $rels[$r->getAttribute('Id')] = $r->getAttribute('Target');
    }
    $sheets = array(); // 並び順 = ワークブックの表示順（1枚目, 2枚目, ...）
    foreach ($wbDoc->getElementsByTagName('sheet') as $sh) {
        $rid = $sh->getAttributeNS('http://schemas.openxmlformats.org/officeDocument/2006/relationships', 'id');
        if ($rid === '') $rid = $sh->getAttribute('r:id');
        if (!isset($rels[$rid])) continue;
        $target = $rels[$rid];
        if ($target[0] !== '/') $target = 'xl/' . preg_replace('#^\./#', '', $target);
        $sheets[] = array($sh->getAttribute('name'), ltrim($target, '/'));
    }

    // 設定に従って対象シートを選ぶ
    $use = array();
    if ($settings['sheet_mode'] === 'all') {
        $use = $sheets;
    } elseif ($settings['sheet_mode'] === 'name') {
        foreach ($sheets as $s) if ($settings['sheet_name'] !== '' && strpos($s[0], $settings['sheet_name']) !== false) $use[] = $s;
        if (!$use) { $zip->close(); return array('rows' => array(), 'error' => '指定名のシートがありません（名前の一部: ' . $settings['sheet_name'] . '）'); }
    } else { // index: 何枚目（初期値1枚目）
        $idx = $settings['sheet_index'] - 1;
        if (isset($sheets[$idx])) $use[] = $sheets[$idx];
        if (!$use) { $zip->close(); return array('rows' => array(), 'error' => 'シートが ' . $settings['sheet_index'] . ' 枚ありません'); }
    }

    $rows = array();
    foreach ($use as $sheet) {
        list($sheetName, $sheetPath) = $sheet;
        $doc = load_dom($zip->getFromName($sheetPath));
        if (!$doc) continue;

        /* --- グリッド化（セル値の解決） --- */
        $grid = array(); // 行番号 => 列文字 => 値
        foreach ($doc->getElementsByTagName('row') as $rowEl) {
            $rowNo = (int)$rowEl->getAttribute('r');
            foreach ($rowEl->getElementsByTagName('c') as $c) {
                $ref = $c->getAttribute('r');
                if (!preg_match('/^([A-Z]+)/', $ref, $m)) continue;
                $col = $m[1];
                $t = $c->getAttribute('t');
                $val = '';
                if ($t === 's') {
                    $v = $c->getElementsByTagName('v')->item(0);
                    if ($v !== null && isset($shared[(int)$v->textContent])) $val = $shared[(int)$v->textContent];
                } elseif ($t === 'inlineStr') {
                    $is = $c->getElementsByTagName('is')->item(0);
                    if ($is !== null) $val = $is->textContent;
                } else {
                    $v = $c->getElementsByTagName('v')->item(0);
                    if ($v !== null) $val = $v->textContent; // 数式セルは計算済みキャッシュ値
                }
                $grid[$rowNo][$col] = trim((string)$val);
            }
        }

        /* --- 結合セルの展開：範囲内の空セルへ左上の値を入れる --- */
        foreach ($doc->getElementsByTagName('mergeCell') as $mc) {
            $ref = $mc->getAttribute('ref');
            if (!preg_match('/^([A-Z]+)([0-9]+)(?::([A-Z]+)([0-9]+))?$/', $ref, $m)) continue;
            $sc = col_to_num($m[1]); $sr = (int)$m[2];
            $ec = isset($m[3]) && $m[3] !== '' ? col_to_num($m[3]) : $sc;
            $er = isset($m[4]) && $m[4] !== '' ? (int)$m[4] : $sr;
            $top = isset($grid[$sr][$m[1]]) ? $grid[$sr][$m[1]] : '';
            if ($top === '') continue;
            for ($r = $sr; $r <= $er; $r++) {
                for ($cn = $sc; $cn <= $ec; $cn++) {
                    $colL = num_to_col($cn);
                    if (!isset($grid[$r][$colL]) || $grid[$r][$colL] === '') $grid[$r][$colL] = $top;
                }
            }
        }

        /* --- 見出し判定＋行抽出（列順に走査、見出しキーは先勝ち＝結合見出し対策） --- */
        ksort($grid);
        $map = null; $headers = array();
        foreach ($grid as $rowNo => $cellsRaw) {
            $cells = $cellsRaw;
            uksort($cells, function($a, $b) { return col_to_num($a) - col_to_num($b); });
            $newMap = array(); $newHeaders = array(); $isHeader = false;
            foreach ($cells as $col => $val) {
                $t = preg_replace('/\s+/u', '', $val);
                if ($t === '') continue;
                $newHeaders[$col] = preg_replace('/\s+/u', ' ', $val);
                // 品番列の見出し判定（自社フォーマットに合わせて、このキーワード並びを編集できます）
                if (!isset($newMap['code']) && (strpos($t, '品番コード') !== false || strpos($t, 'メーカー品番') !== false || strpos($t, '商品品番') !== false || strpos($t, '製品品番') !== false || $t === '品番')) { $newMap['code'] = $col; $isHeader = true; }
                if (!isset($newMap['price']) && (strpos($t, 'プライス品番') !== false || strpos($t, 'タグ品番') !== false)) $newMap['price'] = $col;
                if (!isset($newMap['name']) && strpos($t, '商品名') !== false) $newMap['name'] = $col;
                if (!isset($newMap['group']) && strpos($t, 'グループ番号') !== false) $newMap['group'] = $col;
                if (!isset($newMap['qty']) && strpos($t, '発注数') !== false) $newMap['qty'] = $col;
                if (!isset($newMap['base_price']) && strpos($t, '価格') !== false && strpos($t, '売価') === false) $newMap['base_price'] = $col;
                if (!isset($newMap['cost']) && strpos($t, '推定原価') !== false) $newMap['cost'] = $col;
                if (!isset($newMap['sell']) && strpos($t, '売価') !== false) $newMap['sell'] = $col;
                if (!isset($newMap['spec']) && strpos($t, 'スペック') !== false) $newMap['spec'] = $col;
            }
            if ($isHeader) { $map = $newMap; $headers = $newHeaders; continue; }

            // 合計行は除外
            $firstVal = '';
            foreach ($cells as $v) { if ($v !== '') { $firstVal = $v; break; } }
            if ($firstVal === '合計') continue;

            // 中身のある行だけ索引に入れる（空白の雛形行を除外）
            $meaningful = false;
            foreach ($cells as $v) {
                if ($v === '' || $v === '-') continue;
                if (preg_match('/[0-9]{4,}/', halfwidth_digits($v))) { $meaningful = true; break; }
                if (!is_numeric($v) && mb_strlen($v) >= 2) { $meaningful = true; break; }
            }
            if (!$meaningful) continue;

            $get = function($key) use ($map, $cells) {
                if ($map === null || !isset($map[$key])) return '';
                $col = $map[$key];
                return isset($cells[$col]) ? $cells[$col] : '';
            };
            $price = '';
            if ($map !== null && isset($map['price'])) {
                $pn = col_to_num($map['price']);
                $p1 = fmt_num(isset($cells[num_to_col($pn)]) ? $cells[num_to_col($pn)] : '');
                $p2 = isset($cells[num_to_col($pn + 1)]) ? trim($cells[num_to_col($pn + 1)]) : '';
                $p3 = fmt_num(isset($cells[num_to_col($pn + 2)]) ? $cells[num_to_col($pn + 2)] : '');
                if ($p1 !== '' && $p3 !== '' && $p1 !== $p3) $price = $p1 . '-' . $p3;
                elseif ($p1 !== '' && $p2 !== '' && $p2 !== '-' && $p1 !== $p2) $price = $p1 . $p2;
                else $price = $p1;
            }
            $nonEmpty = array();
            foreach ($cells as $col => $v) if ($v !== '') $nonEmpty[$col] = $v;

            $rows[] = array(
                'sheet' => $sheetName,
                'row' => $rowNo,
                'cells' => $nonEmpty,
                'headers' => $headers,
                'code_col' => ($map !== null && isset($map['code'])) ? $map['code'] : '',
                'fields' => array(
                    'code' => code_key(digits_only(fmt_num($get('code')))),
                    'price' => $price,
                    'name' => $get('name'),
                    'group' => fmt_num($get('group')),
                    'qty' => fmt_num($get('qty')),
                    'base_price' => fmt_num($get('base_price')),
                    'cost' => fmt_num($get('cost')),
                    'sell' => fmt_num($get('sell')),
                    'spec' => $get('spec'),
                ),
            );
        }
    }
    $zip->close();
    return array('rows' => $rows, 'error' => null);
}

/* ---------------- フォルダ走査＋索引 ---------------- */

function scan_files($base, $exts) {
    $out = array();
    $baseReal = realpath($base);
    if ($baseReal === false) return $out;
    $it = new RecursiveIteratorIterator(
        new RecursiveDirectoryIterator($baseReal, FilesystemIterator::SKIP_DOTS),
        RecursiveIteratorIterator::LEAVES_ONLY
    );
    foreach ($it as $f) {
        if (!$f->isFile()) continue;
        $name = $f->getFilename();
        if ($name[0] === '.' || strpos($name, '~$') === 0) continue;
        $ext = strtolower(pathinfo($name, PATHINFO_EXTENSION));
        if (!in_array($ext, $exts, true)) continue;
        $path = $f->getPathname();
        $rel = ltrim(str_replace('\\', '/', substr($path, strlen($baseReal))), '/');
        $out[$rel] = array('path' => $path, 'mtime' => $f->getMTime(), 'size' => $f->getSize());
    }
    ksort($out);
    return $out;
}

function index_signature($settings) {
    return md5(json_encode(array($settings['base_dir'], $settings['sheet_mode'], $settings['sheet_index'], $settings['sheet_name'])));
}

function build_index($base, $exts, $settings, $cacheFile) {
    $files = scan_files($base, $exts);
    $sig = index_signature($settings);
    $cache = array(); $cacheSig = null;
    if (is_file($cacheFile)) {
        $j = json_decode((string)@file_get_contents($cacheFile), true);
        if (is_array($j) && isset($j['files']) && is_array($j['files'])) { $cache = $j['files']; $cacheSig = isset($j['sig']) ? $j['sig'] : null; }
    }
    if ($cacheSig !== $sig) $cache = array(); // シート設定が変わったら作り直し
    $index = array(); $changed = ($cacheSig !== $sig); $errors = array();
    foreach ($files as $rel => $info) {
        $c = isset($cache[$rel]) ? $cache[$rel] : null;
        if ($c && $c['mtime'] === $info['mtime'] && $c['size'] === $info['size']) {
            $index[$rel] = $c;
        } else {
            $r = parse_xlsx($info['path'], $settings);
            $index[$rel] = array('mtime' => $info['mtime'], 'size' => $info['size'], 'rows' => $r['rows'], 'error' => $r['error']);
            $changed = true;
        }
        if (!empty($index[$rel]['error'])) $errors[$rel] = $index[$rel]['error'];
    }
    if (count($cache) !== count($index)) $changed = true;
    if ($changed) @file_put_contents($cacheFile, json_encode(array('sig' => $sig, 'files' => $index), JSON_UNESCAPED_UNICODE), LOCK_EX);
    return array('files' => $files, 'index' => $index, 'errors' => $errors);
}

/* ---------------- ダウンロード（対象フォルダ内のみ許可） ---------------- */

if (isset($_GET['dl'])) {
    $baseReal = realpath($BASE_DIR);
    $target = $baseReal ? realpath($baseReal . '/' . str_replace('\\', '/', (string)$_GET['dl'])) : false;
    $ext = $target ? strtolower(pathinfo($target, PATHINFO_EXTENSION)) : '';
    if ($target && strpos($target, $baseReal) === 0 && is_file($target) && in_array($ext, $EXTS, true)) {
        $name = basename($target);
        header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        header('Content-Length: ' . filesize($target));
        header("Content-Disposition: attachment; filename*=UTF-8''" . rawurlencode($name));
        readfile($target);
        exit;
    }
    http_response_code(404);
    exit('ファイルが見つかりません');
}

/* ---------------- 検索 ---------------- */

$q = isset($_GET['q']) ? trim((string)$_GET['q']) : '';
$results = array(); $fileNameHits = array(); $stats = null; $errors = array();
$searched = ($q !== '');
$baseOk = is_dir($BASE_DIR);
$indexInfo = null;

if ($searched && $baseOk) {
    $built = build_index($BASE_DIR, $EXTS, $SETTINGS, $CACHE_FILE);
    $stats = array('files' => count($built['files']), 'rows' => 0);
    $errors = $built['errors'];
    $qDigits = digits_only($q);
    $qKey = code_key($qDigits);
    $isPriceQuery = (strpos($q, '-') !== false || strpos($q, '－') !== false) && $qDigits !== '';
    $partial = (strlen($qDigits) >= 3 && strlen($qDigits) < 8);

    foreach ($built['index'] as $rel => $entry) {
        $stats['rows'] += count($entry['rows']);
        if ($qDigits !== '' && strpos(digits_only(basename($rel)), $qDigits) !== false) $fileNameHits[] = $rel;
        elseif ($qDigits === '' && mb_stripos(basename($rel), $q) !== false) $fileNameHits[] = $rel;

        foreach ($entry['rows'] as $r) {
            $hit = false;
            // 検索対象列：設定があればその列、無ければ見出し自動判定の品番列（数字）／全列（文字）
            if ($SEARCH_COLS) {
                $targetCols = $SEARCH_COLS;
            } elseif ($qDigits !== '') {
                $targetCols = $r['code_col'] !== '' ? array($r['code_col']) : array_keys($r['cells']);
            } else {
                $targetCols = array_keys($r['cells']);
            }
            if ($isPriceQuery) {
                if ($r['fields']['price'] !== '' && digits_only($r['fields']['price']) === $qDigits) $hit = true;
            }
            if (!$hit && $qDigits !== '') {
                foreach ($targetCols as $col) {
                    if (!isset($r['cells'][$col])) continue;
                    foreach (number_tokens($r['cells'][$col]) as $tok) {
                        if (!$partial && code_key($tok) === $qKey && strlen($tok) >= 4) { $hit = true; break 2; }
                        if ($partial && strpos(code_key($tok), $qDigits) !== false) { $hit = true; break 2; }
                    }
                }
            }
            if (!$hit && $qDigits === '') {
                foreach ($targetCols as $col) {
                    if (isset($r['cells'][$col]) && mb_stripos($r['cells'][$col], $q) !== false) { $hit = true; break; }
                }
            }
            if ($hit) { $r['file'] = $rel; $results[] = $r; }
        }
    }
}

// 索引の現況（設定パネル表示用）
if (is_file($CACHE_FILE)) {
    $cj = json_decode((string)@file_get_contents($CACHE_FILE), true);
    if (is_array($cj) && isset($cj['files'])) {
        $n = 0; foreach ($cj['files'] as $e) $n += count($e['rows']);
        $indexInfo = array('files' => count($cj['files']), 'rows' => $n, 'mtime' => filemtime($CACHE_FILE), 'size' => filesize($CACHE_FILE));
    }
}

/* 表示列（カスタム時）の見出しラベル：最初のヒットの見出しから取る */
$displayLabels = array();
if ($DISPLAY_COLS && $results) {
    foreach ($DISPLAY_COLS as $col) {
        $label = $col . '列';
        foreach ($results as $r) {
            if (!empty($r['headers'][$col])) { $label = $r['headers'][$col]; break; }
        }
        $displayLabels[$col] = $label;
    }
}
function display_value($r, $col) {
    // プライス品番の見出し列は F/G/H をつないだ結合表記を優先
    if (!empty($r['headers'][$col])) {
        $t = preg_replace('/\s+/u', '', $r['headers'][$col]);
        if (strpos($t, 'プライス品番') !== false || strpos($t, 'タグ品番') !== false) {
            if ($r['fields']['price'] !== '') return $r['fields']['price'];
        }
        if (strpos($t, '品番コード') !== false || strpos($t, 'メーカー品番') !== false) {
            if ($r['fields']['code'] !== '') return $r['fields']['code'];
        }
    }
    return isset($r['cells'][$col]) ? $r['cells'][$col] : '';
}
?>
<!DOCTYPE html>
<html lang="ja">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Excel横断検索（Excel_Multi-Sheet_Search）｜品番コード → プライス品番</title>
<style>
  body{font-family:"Hiragino Kaku Gothic ProN","Yu Gothic",sans-serif;margin:0;background:#f4f5f7;color:#222}
  .wrap{max-width:1080px;margin:0 auto;padding:20px 14px 48px}
  h1{font-size:20px;margin:8px 0 4px}
  .sub{color:#555;font-size:13px;margin-bottom:16px}
  form.search{display:flex;gap:8px;flex-wrap:wrap}
  form.search input[type=text]{flex:1;min-width:220px;font-size:18px;padding:10px 12px;border:1px solid #bbb;border-radius:8px}
  form.search button{font-size:16px;padding:10px 22px;border:0;border-radius:8px;background:#1f6f43;color:#fff;cursor:pointer}
  .hint{font-size:12.5px;color:#555;margin-top:8px}
  .stats{font-size:13px;color:#444;margin:18px 0 8px}
  table{width:100%;border-collapse:collapse;background:#fff;font-size:14px}
  th,td{border:1px solid #ddd;padding:7px 8px;text-align:left;vertical-align:top}
  th{background:#eef3f0;white-space:nowrap}
  td.price{font-weight:700;font-size:16px;white-space:nowrap}
  td.num{text-align:right;white-space:nowrap}
  .file a{color:#1f6f43}
  .loc{color:#777;font-size:12px}
  .none,.warn{background:#fff;border:1px solid #ddd;border-radius:8px;padding:16px}
  .warn{border-color:#e5c07b;background:#fffaf0;margin-bottom:14px}
  .err{background:#fff4f4;border:1px solid #e5b4b4;border-radius:8px;padding:10px 14px;font-size:13px;margin-top:14px}
  details.settings{margin-top:28px;background:#fff;border:1px solid #ddd;border-radius:8px;padding:12px 16px}
  details.settings summary{cursor:pointer;font-weight:700}
  .sgrid{display:grid;grid-template-columns:150px 1fr;gap:8px 12px;margin-top:12px;font-size:14px;align-items:center}
  .sgrid input[type=text],.sgrid input[type=number],.sgrid select{padding:7px 9px;border:1px solid #bbb;border-radius:6px;font-size:14px;width:100%;box-sizing:border-box}
  .sgrid .note{grid-column:2;color:#777;font-size:12px}
  .sbtn{margin-top:12px;display:flex;gap:8px;flex-wrap:wrap}
  .sbtn button{padding:8px 18px;border:0;border-radius:7px;background:#1f6f43;color:#fff;cursor:pointer}
  .sbtn button.sub2{background:#666}
  @media(max-width:760px){ .hideS{display:none} .sgrid{grid-template-columns:1fr} .sgrid .note{grid-column:1} }
</style>
</head>
<body>
<div class="wrap">
  <h1>Excel横断検索（Excel_Multi-Sheet_Search）</h1>
  <div class="sub">品番コードから「プライス品番・タグ品番（999-XXXX）」を探します。対象フォルダ＋子フォルダ内の全Excelが対象です（日本語フォルダ・日本語ファイル名対応）。</div>

  <?php if (!$baseOk): ?>
  <div class="warn"><b>検索フォルダが見つかりません：</b><?php echo h($BASE_DIR); ?><br>
  Docker のボリューム指定（例：<code>/Work</code>）か、下の設定の「検索フォルダ」を確認してください。</div>
  <?php endif; ?>
  <?php if (isset($_GET['saved'])): ?><div class="stats">設定を保存しました（索引は次回検索時に自動で作り直されます）。</div><?php endif; ?>
  <?php if ($displayTrunc): ?><div class="warn">表示列は最大10個までです。先頭の10個を適用しています。</div><?php endif; ?>

  <form class="search" method="get" action="">
    <input type="text" name="q" value="<?php echo h($q); ?>" placeholder="品番コード（8桁）例: 12345678" inputmode="numeric" autofocus>
    <button type="submit">検索</button>
  </form>
  <div class="hint">入力例：品番コード <b>12345678</b>（8桁・完全一致）／ 途中までの数字でも部分一致 ／ <b>999-0001</b> プライス品番からも可 ／ 商品名でも可 ／ セル内に「12345678+12345679」と複数入っている場合も個別にヒットします</div>

<?php if ($searched): ?>
  <?php if ($baseOk): ?>
  <div class="stats">
    検索対象：Excel <?php echo (int)$stats['files']; ?> ファイル ／ 索引行 <?php echo (int)$stats['rows']; ?> 件
    → ヒット <b><?php echo count($results); ?></b> 件
  </div>

  <?php if ($results): ?>
  <table>
    <tr>
      <?php if ($DISPLAY_COLS): ?>
        <?php foreach ($DISPLAY_COLS as $col): ?><th><?php echo h($displayLabels[$col]); ?><br><span class="loc"><?php echo h($col); ?>列</span></th><?php endforeach; ?>
      <?php else: ?>
        <th>プライス品番<br>タグ品番</th><th>品番コード</th><th>商品名</th>
        <th class="hideS">グループ</th><th>推定売価</th><th class="hideS">基準価格</th>
        <th class="hideS">推定原価</th><th class="hideS">スペック</th>
      <?php endif; ?>
      <th>Excel名</th>
    </tr>
    <?php foreach ($results as $r): ?>
    <tr>
      <?php if ($DISPLAY_COLS): ?>
        <?php foreach ($DISPLAY_COLS as $col): ?><td><?php echo h(display_value($r, $col)); ?></td><?php endforeach; ?>
      <?php else: ?>
        <td class="price"><?php echo h($r['fields']['price']); ?></td>
        <td><?php echo h($r['fields']['code']); ?></td>
        <td><?php echo h($r['fields']['name']); ?></td>
        <td class="hideS"><?php echo h($r['fields']['group']); ?></td>
        <td class="num"><?php echo h(fmt_money($r['fields']['sell'])); ?></td>
        <td class="num hideS"><?php echo h(fmt_money($r['fields']['base_price'])); ?></td>
        <td class="num hideS"><?php echo h(fmt_money($r['fields']['cost'])); ?></td>
        <td class="hideS"><?php echo h($r['fields']['spec']); ?></td>
      <?php endif; ?>
      <td class="file">
        <a href="?dl=<?php echo rawurlencode($r['file']); ?>"><?php echo h(basename($r['file'])); ?></a>
        <div class="loc"><?php echo h(dirname($r['file']) === '.' ? '（親フォルダ）' : dirname($r['file'])); ?><br>
        <?php echo h($r['sheet']); ?> ／ <?php echo (int)$r['row']; ?>行目</div>
      </td>
    </tr>
    <?php endforeach; ?>
  </table>
  <?php else: ?>
  <div class="none">「<?php echo h($q); ?>」に一致する行は見つかりませんでした。<br>
  <span class="loc">数字の打ち間違い、対象フォルダ外のExcel、読み込むシート設定（現在：<?php echo h($SETTINGS['sheet_mode'] === 'index' ? $SETTINGS['sheet_index'] . '枚目' : ($SETTINGS['sheet_mode'] === 'name' ? '名前「' . $SETTINGS['sheet_name'] . '」を含む' : '全シート')); ?>）や検索列設定が違う場合はヒットしません。</span></div>
  <?php endif; ?>

  <?php if ($fileNameHits): ?>
  <div class="stats">ファイル名に「<?php echo h($q); ?>」を含むExcel：</div>
  <table>
    <?php foreach ($fileNameHits as $rel): ?>
    <tr><td class="file"><a href="?dl=<?php echo rawurlencode($rel); ?>"><?php echo h(basename($rel)); ?></a>
    <span class="loc">／ <?php echo h(dirname($rel) === '.' ? '（親フォルダ）' : dirname($rel)); ?></span></td></tr>
    <?php endforeach; ?>
  </table>
  <?php endif; ?>

  <?php if ($errors): ?>
  <div class="err"><b>読み込めなかったファイル</b><br>
  <?php foreach ($errors as $rel => $msg): ?>
    <?php echo h($rel); ?>：<?php echo h($msg); ?><br>
  <?php endforeach; ?>
  </div>
  <?php endif; ?>
  <?php endif; ?>
<?php endif; ?>

  <details class="settings" <?php echo isset($_GET['saved']) ? 'open' : ''; ?>>
    <summary>設定（シート・検索列・表示列）／索引の状態</summary>
    <form method="post" action="">
      <input type="hidden" name="action" value="save_settings">
      <div class="sgrid">
        <label>検索フォルダ</label>
        <input type="text" name="base_dir" value="<?php echo h($SETTINGS['base_dir']); ?>">
        <div class="note">Dockerでは通常 <code>/Work</code>（ボリュームで渡したフォルダ）。日本語フォルダ可。この配下の子フォルダも全て検索します。</div>

        <label>読み込むシート</label>
        <div>
          <select name="sheet_mode" style="width:auto">
            <option value="index" <?php echo $SETTINGS['sheet_mode'] === 'index' ? 'selected' : ''; ?>>何枚目かで指定</option>
            <option value="name" <?php echo $SETTINGS['sheet_mode'] === 'name' ? 'selected' : ''; ?>>シート名の一部で指定</option>
            <option value="all" <?php echo $SETTINGS['sheet_mode'] === 'all' ? 'selected' : ''; ?>>全シート</option>
          </select>
          <input type="number" name="sheet_index" value="<?php echo (int)$SETTINGS['sheet_index']; ?>" min="1" style="width:70px"> 枚目
          ／ 名前の一部：<input type="text" name="sheet_name" value="<?php echo h($SETTINGS['sheet_name']); ?>" style="width:160px" placeholder="例: 基準価格">
        </div>
        <div class="note">初期値は「1枚目」。シート名がファイルごとに変わる場合は「何枚目」が確実です。</div>

        <label>検索列（A〜Z）</label>
        <input type="text" name="search_cols" value="<?php echo h($SETTINGS['search_cols']); ?>" placeholder="例: N ／ A-Z ／ B,N,P（空欄＝見出しから自動判定）">
        <div class="note">入力した品番を、この列のセルから探します。セル内に「12345678+12345679」のように複数入っていても個別に照合します。</div>

        <label>表示列（最大10）</label>
        <input type="text" name="display_cols" value="<?php echo h($SETTINGS['display_cols']); ?>" placeholder="例: F,N,C,L,O（空欄＝標準表示）">
        <div class="note">結果表に出す列を列文字で指定します（最大10個）。見出し行がある列は見出し名で表示されます。</div>
      </div>
      <div class="sbtn">
        <button type="submit">設定を保存</button>
      </div>
    </form>
    <form method="post" action="" onsubmit="return confirm('索引を作り直します。次回検索時に全ファイルを読み直します。よろしいですか？');">
      <input type="hidden" name="action" value="rebuild">
      <div class="sbtn">
        <button type="submit" class="sub2">索引を作り直す</button>
      </div>
    </form>
    <div class="stats">
      索引の保存先：<?php echo h($INDEX_DIR); ?>（search_index_v2.json）<br>
      <?php if ($indexInfo): ?>
      現在の索引：Excel <?php echo (int)$indexInfo['files']; ?> ファイル ／ <?php echo (int)$indexInfo['rows']; ?> 行 ／ <?php echo number_format($indexInfo['size']); ?> バイト ／ 最終更新 <?php echo date('Y-m-d H:i', $indexInfo['mtime']); ?>
      <?php else: ?>
      索引はまだ作られていません（初回検索時に自動作成されます）。
      <?php endif; ?>
    </div>
  </details>
</div>
</body>
</html>
