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
$INDEX_MOUNTED = false;
$mountsTxt = @file_get_contents('/proc/mounts');
if ($mountsTxt) $INDEX_MOUNTED = (bool)preg_match('/\s\/Index\s/', $mountsTxt);
$INDEX_FALLBACK = ($INDEX_DIR !== '/Index');
$IDX_WARN = is_dir('/Index') && ($INDEX_FALLBACK || !$INDEX_MOUNTED);
// 以前コンテナ内に保存されていた設定・索引があれば /Index へ自動移行する
if ($INDEX_DIR === '/Index' && __DIR__ !== '/Index') {
    foreach (array('settings.json', 'search_index_v2.json') as $fn) {
        if (!is_file($INDEX_DIR . '/' . $fn) && is_file(__DIR__ . '/' . $fn)) @copy(__DIR__ . '/' . $fn, $INDEX_DIR . '/' . $fn);
    }
}

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
        'bg_mode'      => 'on',  // on / off / light / black
        'bg_nodes'     => 'mid', // low=44 / mid=88 / high=160
        'layout'       => '',      // 表示レイアウトJSON（空欄=標準表示）
        'image_tpl'    => '',      // 画像URLテンプレート（空欄=画像なし）
        'show_age'     => '0',     // ファイルの時期表示 1=する 0=しない
        'nodes_show'   => '1',     // ノード表示 1=する 0=完全に非表示
    );
}
function load_settings($file) {
    $s = default_settings();
    if (is_file($file)) {
        $j = json_decode((string)@file_get_contents($file), true);
        if (is_array($j)) $s = array_merge($s, $j); // 未知のキーも保持する（版を跨いでも設定が消えないように）
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
        $SETTINGS['bg_mode'] = (string)($_POST['bg_mode'] ?? 'on');
        $SETTINGS['bg_nodes'] = (string)($_POST['bg_nodes'] ?? 'mid');
        $SETTINGS['layout'] = trim((string)($_POST['layout'] ?? ''));
        $SETTINGS['image_tpl'] = trim((string)($_POST['image_tpl'] ?? ''));
        $SETTINGS['show_age'] = (($_POST['show_age'] ?? '0') === '1') ? '1' : '0';
        $SETTINGS['nodes_show'] = (($_POST['nodes_show'] ?? '1') === '0') ? '0' : '1';
        if (!in_array($SETTINGS['sheet_mode'], array('index', 'name', 'all'), true)) $SETTINGS['sheet_mode'] = 'index';
        if (!in_array($SETTINGS['bg_mode'], array('on', 'off', 'light', 'black'), true)) $SETTINGS['bg_mode'] = 'on';
        if (!in_array($SETTINGS['bg_nodes'], array('low', 'mid', 'high'), true)) $SETTINGS['bg_nodes'] = 'mid';
        if (is_file($SETTINGS_FILE)) @copy($SETTINGS_FILE, $SETTINGS_FILE . '.bak');
        @file_put_contents($SETTINGS_FILE, json_encode($SETTINGS, JSON_UNESCAPED_UNICODE | JSON_PRETTY_PRINT), LOCK_EX);
    } elseif ($_POST['action'] === 'rebuild') {
        @unlink($CACHE_FILE);
    }
    $loc = strtok($_SERVER['REQUEST_URI'], '?') . '?saved=1';
    header('Location: ' . $loc);
    exit;
}

$BASE_DIR = $SETTINGS['base_dir'];
$BG_MODE = $SETTINGS['bg_mode'] ?? 'on';
if (!in_array($BG_MODE, array('on', 'off', 'light', 'black'), true)) $BG_MODE = 'on';
$BG_NODES = $SETTINGS['bg_nodes'] ?? 'mid';
if (!in_array($BG_NODES, array('low', 'mid', 'high'), true)) $BG_NODES = 'mid';
$BG_COUNT = array('low' => 44, 'mid' => 88, 'high' => 160)[$BG_NODES];

/* ---------------- 小物関数 ---------------- */

define('APP_VERSION', '1.18');
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
$KEY_LABELS = array('price' => 'プライス品番・タグ品番', 'code' => '品番コード', 'name' => '商品名', 'group' => 'グループ', 'sell' => '推定売価', 'base_price' => '基準価格', 'cost' => '推定原価', 'spec' => 'スペック', 'image' => '画像', 'file' => 'Excel名');
/* 表示レイアウト（JSON）: [{"key":"price","label":"..."},{"col":"N","label":"..."}] */
function parse_layout($json) {
    if (trim((string)$json) === '') return null;
    $a = json_decode($json, true);
    if (!is_array($a)) return null;
    $out = array();
    foreach ($a as $e) {
        if (!is_array($e)) continue;
        $key = isset($e['key']) ? (string)$e['key'] : '';
        $col = isset($e['col']) ? strtoupper(trim((string)$e['col'])) : '';
        $label = isset($e['label']) ? (string)$e['label'] : '';
        if (in_array($key, array('price','code','name','group','sell','base_price','cost','spec','image','file'), true)) {
            $out[] = array('key' => $key, 'col' => '', 'label' => $label);
        } elseif (preg_match('/^[A-Z]{1,2}$/', $col)) {
            $out[] = array('key' => '', 'col' => $col, 'label' => $label);
        }
        if (count($out) >= 15) break;
    }
    return $out ? $out : null;
}
function layout_label($e, $displayLabels) {
    global $KEY_LABELS;
    if ($e['label'] !== '') return $e['label'];
    if ($e['key'] !== '') return $KEY_LABELS[$e['key']] ?? $e['key'];
    return $displayLabels[$e['col']] ?? ($e['col'] . '列');
}
function layout_value($r, $e) {
    if ($e['key'] !== '') {
        switch ($e['key']) {
            case 'price': return array($r['fields']['price'], 'price');
            case 'code': return array($r['fields']['code'], '');
            case 'name': return array($r['fields']['name'], '');
            case 'group': return array($r['fields']['group'], 'hideS');
            case 'sell': return array(fmt_money($r['fields']['sell']), 'num');
            case 'base_price': return array(fmt_money($r['fields']['base_price']), 'num hideS');
            case 'cost': return array(fmt_money($r['fields']['cost']), 'num hideS');
            case 'spec': return array($r['fields']['spec'], 'hideS');
        }
    }
    return array(display_value($r, $e['col']), '');
}
function file_age_label($ts) {
    if (!$ts) return '';
    $d = floor((time() - $ts) / 86400);
    if ($d < 1) return '今日のファイル';
    if ($d < 7) return $d . '日前のファイル';
    if ($d < 30) return '約' . max(1, (int)round($d / 7)) . '週間前のファイル';
    $base = date('Y年n月', $ts) . '頃のファイル';
    if ($d < 365) return $base . '（約' . max(1, (int)round($d / 30.4)) . 'ヶ月前）';
    return $base;
}
function age_div($r) {
    global $BASE_DIR, $AGE_ON;
    if (!$AGE_ON) return '';
    static $cache = array();
    $f = $r['file'];
    if (!array_key_exists($f, $cache)) $cache[$f] = @filemtime($BASE_DIR . '/' . $f);
    $label = file_age_label($cache[$f] ? $cache[$f] : null);
    return $label === '' ? '' : '<div class="loc">' . h($label) . '</div>';
}
function image_url($r, $tpl) {
    $tpl = trim((string)$tpl);
    if ($tpl === '') return '';
    $map = array(
        '{code}'  => rawurlencode($r['fields']['code']),
        '{price}' => rawurlencode($r['fields']['price']),
        '{name}'  => rawurlencode($r['fields']['name']),
        '{group}' => rawurlencode($r['fields']['group'])
    );
    return strtr($tpl, $map);
}
function img_cell($r, $tpl) {
    $u = image_url($r, $tpl);
    if ($u === '') return '';
    $u = h($u);
    return '<img src="' . $u . '" alt="" loading="lazy" style="max-width:72px;max-height:72px;border-radius:6px;cursor:zoom-in" onclick="openImgPopup(this.src)" onerror="imgMissing(this)">';
}
$IMG_ON = trim((string)($SETTINGS['image_tpl'] ?? '')) !== '';
$AGE_ON = ($SETTINGS['show_age'] ?? '0') === '1';
$NODES_SHOW = ($SETTINGS['nodes_show'] ?? '1') === '1';
$LAYOUT = parse_layout($SETTINGS['layout'] ?? '');
$LAYOUT_EXAMPLE = json_encode(array(
    array('key' => 'price', 'label' => $KEY_LABELS['price']),
    array('key' => 'code', 'label' => $KEY_LABELS['code']),
    array('key' => 'name', 'label' => $KEY_LABELS['name']),
    array('key' => 'group', 'label' => $KEY_LABELS['group']),
    array('key' => 'sell', 'label' => $KEY_LABELS['sell']),
    array('key' => 'base_price', 'label' => $KEY_LABELS['base_price']),
    array('key' => 'cost', 'label' => $KEY_LABELS['cost']),
    array('key' => 'spec', 'label' => $KEY_LABELS['spec']),
    array('key' => 'file', 'label' => $KEY_LABELS['file'])
), JSON_UNESCAPED_UNICODE);
$LAYOUT_EFFECTIVE = $LAYOUT;
if (!$LAYOUT_EFFECTIVE) {
    $LAYOUT_EFFECTIVE = array();
    foreach (array('price','code','name','group','sell','base_price','cost','spec','file') as $k) {
        $LAYOUT_EFFECTIVE[] = array('key' => $k, 'col' => '', 'label' => $KEY_LABELS[$k]);
    }
}


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
        foreach ($sheets as $i => $s) $use[] = array($s[0], $s[1], $i + 1);
    } elseif ($settings['sheet_mode'] === 'name') {
        foreach ($sheets as $i => $s) if ($settings['sheet_name'] !== '' && strpos($s[0], $settings['sheet_name']) !== false) $use[] = array($s[0], $s[1], $i + 1);
        if (!$use) { $zip->close(); return array('rows' => array(), 'error' => '指定名のシートがありません（名前の一部: ' . $settings['sheet_name'] . '）'); }
    } else { // index: 何枚目（初期値1枚目）
        $idx = $settings['sheet_index'] - 1;
        if (isset($sheets[$idx])) $use[] = array($sheets[$idx][0], $sheets[$idx][1], $idx + 1);
        if (!$use) { $zip->close(); return array('rows' => array(), 'error' => 'シートが ' . $settings['sheet_index'] . ' 枚ありません'); }
    }

    $rows = array();
    foreach ($use as $sheet) {
        list($sheetName, $sheetPath, $sheetNo) = $sheet;
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
                'sno' => $sheetNo,
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
    try {
        $it = new RecursiveIteratorIterator(
            new RecursiveDirectoryIterator($baseReal, FilesystemIterator::SKIP_DOTS),
            RecursiveIteratorIterator::LEAVES_ONLY
        );
    } catch (Exception $e) {
        // 読めないフォルダ（権限不足など）では500にせず、警告として返す
        $GLOBALS['SCAN_ERROR'] = '検索フォルダを読み込めませんでした。フォルダの読み取り権限（コンテナの www-data が読めること）を確認してください：' . $base;
        return $out;
    }
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
    return md5(json_encode(array($settings['base_dir'], $settings['sheet_mode'], $settings['sheet_index'], $settings['sheet_name'], 'fmt3')));
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
    if (!empty($GLOBALS['SCAN_ERROR'])) $errors['（検索フォルダ）'] = $GLOBALS['SCAN_ERROR'];
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
$baseOk = is_dir($BASE_DIR) && is_readable($BASE_DIR);
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
  @media(max-width:760px){
    .sgrid{grid-template-columns:1fr} .sgrid .note{grid-column:1}
    table.rlist, table.rlist tbody, table.rlist tr, table.rlist td{display:block;width:auto}
    table.rlist tr.rhead{display:none}
    table.rlist tr{border:1px solid #c9c9c9;border-radius:10px;margin:10px 0;padding:8px 10px;background:rgba(255,255,255,.65)}
    body.dark table.rlist tr{border-color:#5a4632;background:rgba(30,22,14,.65)}
    table.rlist td{border:0;border-bottom:1px dashed rgba(0,0,0,.1);padding:4px 0;text-align:left}
    table.rlist td:last-child{border-bottom:0}
    table.rlist td::before{content:attr(data-label);display:block;font-size:11px;font-weight:600;color:#888}
    table.rlist td.price{font-size:16px}
  }

  #bgcanvas{position:fixed;inset:0;z-index:0;pointer-events:none}
  .wrap{position:relative;z-index:1}
  body.dark{background:#16120d;color:#e9dfcd}
  body.dark .sub,body.dark .hint{color:#b8a888}
  body.dark .stats{color:#cbbfa5}
  body.dark table{background:#211a12}
  body.dark th,body.dark td{border-color:#4a3a22}
  body.dark th{background:#2c2317}
  body.dark .file a{color:#ff9e3d}
  body.dark .loc{color:#a08c66}
  body.dark .none,body.dark details.settings{background:#211a12;border-color:#4a3a22}
  body.dark .warn{background:#2a2113;border-color:#8a6d1f}
  body.dark .err{background:#2c1717;border-color:#7a3a3a}
  body.dark input[type=text],body.dark input[type=number],body.dark select{background:#171310;color:#e9dfcd;border-color:#6b5322}
  body.dark form.search button,body.dark .sbtn button{background:#b35c00}
  body.dark .sbtn button.sub2{background:#555}
  body.dark a{color:#ff9e3d}

  .modal-bg{position:fixed;inset:0;background:rgba(0,0,0,.45);z-index:60;display:none}
  .modal{background:#fff;border-radius:10px;max-width:660px;margin:5vh auto;padding:16px 18px;max-height:86vh;overflow:auto;box-shadow:0 8px 30px rgba(0,0,0,.3)}
  body.dark .modal{background:#211a12}
  .lrow{display:flex;gap:6px;align-items:center;margin-bottom:6px;flex-wrap:wrap}
  .lrow select,.lrow input{padding:6px 8px;border:1px solid #bbb;border-radius:6px;font-size:13.5px}
  .edbtn{padding:7px 13px;border:0;border-radius:7px;background:#666;color:#fff;cursor:pointer;font-size:13.5px}
  .edbtn.mv{padding:5px 9px}
  body.dark a{color:#f0a35e}
  .file details{margin-top:1px}
  .file summary{cursor:pointer;color:#888;font-size:12px}
</style>
<script>
function imgMissing(img){
  var t = (img.parentNode && img.parentNode.tagName === 'A') ? img.parentNode : img;
  var sp = document.createElement('span');
  sp.className = 'loc'; sp.textContent = 'ありません';
  t.replaceWith(sp);
}
</script>
</head>
<body<?php echo $BG_MODE === 'black' ? ' class="dark"' : ''; ?>>
<?php if ($BG_MODE !== 'off' && $NODES_SHOW): ?>
<canvas id="bgcanvas"></canvas>
<script>
var BG_MODE = '<?php echo $BG_MODE; ?>', BG_N = <?php echo (int)$BG_COUNT; ?>;
var BG_PALS = {
  on:    {n:'100,116,139', l:'100,116,139', na:0.55, la:0.30},
  light: {n:'37,99,235',   l:'37,99,235',   na:0.60, la:0.34},
  black: {n:'255,150,40',  l:'255,120,20',  na:0.80, la:0.42}
};
var bgRun = false, bgCv = null, bgCtx = null, bgNS = [], bgW = 0, bgH = 0;
function bgResize(){ if (bgCv) { bgW = bgCv.width = window.innerWidth; bgH = bgCv.height = window.innerHeight; } }
window.addEventListener('resize', bgResize);
function bgTick(){
  if (!bgRun) return;
  var P = BG_PALS[BG_MODE] || BG_PALS.on, LINK = 130;
  bgCtx.clearRect(0, 0, bgW, bgH);
  var i, j, a, b, dx, dy, d;
  for (i = 0; i < bgNS.length; i++) {
    a = bgNS[i]; a.x += a.vx; a.y += a.vy;
    if (a.x < 0) a.x += bgW; if (a.x > bgW) a.x -= bgW;
    if (a.y < 0) a.y += bgH; if (a.y > bgH) a.y -= bgH;
  }
  for (i = 0; i < bgNS.length; i++) for (j = i + 1; j < bgNS.length; j++) {
    a = bgNS[i]; b = bgNS[j];
    dx = a.x - b.x; dy = a.y - b.y; d = Math.sqrt(dx*dx + dy*dy);
    if (d < LINK) {
      bgCtx.strokeStyle = 'rgba(' + P.l + ',' + (P.la * (1 - d / LINK)).toFixed(3) + ')';
      bgCtx.lineWidth = 1;
      bgCtx.beginPath(); bgCtx.moveTo(a.x, a.y); bgCtx.lineTo(b.x, b.y); bgCtx.stroke();
    }
  }
  for (i = 0; i < bgNS.length; i++) {
    a = bgNS[i];
    bgCtx.fillStyle = a.g ? 'rgba(110,220,130,' + P.na + ')' : 'rgba(' + P.n + ',' + P.na + ')';
    bgCtx.beginPath(); bgCtx.arc(a.x, a.y, 1.6, 0, 6.2832); bgCtx.fill();
  }
  requestAnimationFrame(bgTick);
}
function bgStart(){
  bgCv = document.getElementById('bgcanvas');
  if (!bgCv || bgRun) return;
  bgCtx = bgCv.getContext('2d'); bgResize();
  if (!bgNS.length) {
    for (var i = 0; i < BG_N; i++) bgNS.push({x: Math.random()*bgW, y: Math.random()*bgH, vx: (Math.random()-0.5)*0.35, vy: (Math.random()-0.5)*0.35, g: (BG_MODE === 'black' && i % 6 === 0)});
  }
  bgRun = true; requestAnimationFrame(bgTick);
}
bgStart();
</script>
<?php endif; ?>
<?php
// ===== 権限チェック（警告内とポップアップで共用するため、HTMLを先に組み立てる） =====
ob_start();
  // ===== 権限チェック（実際に読めるか・書けるかをその場でテストする） =====
  $envPuid = getenv('PUID'); $envPgid = getenv('PGID');
  $runName = '?'; $runUid = '?'; $runGid = '?';
  if (function_exists('posix_getpwuid') && function_exists('posix_geteuid')) {
      $pw = @posix_getpwuid(posix_geteuid());
      if ($pw) { $runName = $pw['name']; $runUid = $pw['uid']; $runGid = $pw['gid']; }
  }
  $workList = @scandir($BASE_DIR);
  $workReadable = ($workList !== false);
  $idxWritable = false; $idxNote = '';
  if (is_dir($INDEX_DIR)) {
      $tmpf = @tempnam($INDEX_DIR, '.permtest_');
      if ($tmpf !== false) { $idxWritable = true; @unlink($tmpf); }
      else $idxNote = '（tempnam失敗）';
  }
  $okNg = function ($b) { return $b ? '<b style="color:#2e7d32">OK</b>' : '<b style="color:#c62828">NG</b>'; };
  echo '<div style="margin-top:8px"><b>権限チェック</b><ul style="margin:4px 0 0 18px;padding:0">';
  echo '<li>環境変数 PUID: ' . ($envPuid === false || $envPuid === '' ? '（未設定）' : h((string)$envPuid)) . ' / PGID: ' . ($envPgid === false || $envPgid === '' ? '（未設定）' : h((string)$envPgid)) . '</li>';
  echo '<li>実行ユーザー: ' . h($runName) . '（UID: ' . h((string)$runUid) . ' / GID: ' . h((string)$runGid) . '）</li>';
  echo '<li>検索フォルダ ' . h($BASE_DIR) . ' の読み取り: ' . $okNg($workReadable) . ($workReadable ? '（中身 ' . count($workList) . ' 件）' : '') . '</li>';
  echo '<li>索引フォルダ ' . h($INDEX_DIR) . ' の書き込み: ' . $okNg($idxWritable) . h($idxNote) . '</li>';
  echo '<li>/Index マウント: ' . ($INDEX_MOUNTED ? 'あり' : '<b style="color:#c62828">なし（ボリューム未割り当ての可能性）</b>') . '</li>';
  echo '<li>/Index 書込可否: ' . (is_dir('/Index') ? (is_writable('/Index') ? 'OK' : '<b style="color:#c62828">NG（書けません）</b>') : 'フォルダなし') . '</li>';
  echo '</ul>';
  if ($envPuid !== false && $envPuid !== '' && (string)$runUid !== (string)$envPuid) {
      echo '<span class="muted">※ PUID は指定されていますが実行UIDに反映されていません。コンテナが再作成されていないか、古いイメージの可能性があります（イメージの更新＋作り直しをしてください）。</span><br>';
  }
  if ($envPgid !== false && $envPgid !== '' && (string)$runGid !== (string)$envPgid) {
      echo '<span class="muted">※ PGID は指定されていますが実行GIDに反映されていません。イメージ v1.3 以降で解消します。</span><br>';
  }
  if (is_dir('/Index') && !is_writable('/Index')) {
      echo '<span class="muted">※ /Index に書けません。フォルダの所有者が「33」（www-data）と表示されるのは正常です（NASに存在しない番号のUIDでも動作します）。原因は所有者ではなく、フォルダのACLに書き込み許可のエントリが無いことです。File Station で Index フォルダの「権限」タブを開き、everyone に読み取り＋書き込みを許可してください（適用先：このフォルダ、子フォルダ、ファイル）。</span><br>';
  }
  if ($workReadable === false && (string)$runUid === (string)@fileowner($BASE_DIR)) {
      echo '<span class="muted">※ UIDは所有者と一致しているのに読めません。Synology の ACL 管理フォルダ（権限表示 0000）の可能性があります。その場合は File Station の権限で「everyone」に読み取りを付けるのが確実です。</span><br>';
  }
  echo '</div>';
$PERM_HTML = ob_get_clean();
?>
<div class="wrap">
  <h1>Excel横断検索（Excel_Multi-Sheet_Search）</h1>
  <div class="sub">品番コードから「プライス品番・タグ品番（999-XXXX）」を探します。対象フォルダ＋子フォルダ内の全Excelが対象です（日本語フォルダ・日本語ファイル名対応）。</div>

  <?php if (!$baseOk || $IDX_WARN): ?>
  <div class="warn"><?php if (!$baseOk): ?><b>検索フォルダが見つからないか、読み取り権限がありません：</b><?php echo h($BASE_DIR); ?><br>
  Docker のボリューム指定（例：<code>/Work</code>）か、下の設定の「検索フォルダ」を確認してください。<br><?php else: ?><b>索引フォルダ（/Index）に書き込めないか、ボリュームが割り当たっていません。</b>設定と索引がコンテナ内に保存されるため、更新・リセットで消えます。<br><?php endif; ?>
  <?php
  $diag = array();
  $diag[] = 'フォルダ存在: ' . (is_dir($BASE_DIR) ? 'あり' : 'なし');
  if (is_dir($BASE_DIR)) {
      $diag[] = '権限: ' . substr(sprintf('%o', @fileperms($BASE_DIR)), -4);
      $diag[] = '所有者UID: ' . @fileowner($BASE_DIR);
  }
  if (function_exists('posix_geteuid')) $diag[] = '実行中UID: ' . posix_geteuid();
  echo '<span class="muted">診断: ' . h(implode(' / ', $diag)) . '</span><br>';
  ?>
  <span class="muted">フォルダの所有者と実行中UIDが違う場合は、コンテナの環境変数 PUID / PGID に所有者のUID/GID（例：1026 / 100）を指定すると読めるようになります。</span><br>
  <span class="muted">※ Synology 等の ACL 管理フォルダでは、対象フォルダ自体に閲覧権限を付ける方が確実です。File Station の権限設定で「everyone」に読み取り（一覧表示／データの読取り＋フォルダ間の移動／実行ファイル＋属性の読み取り）を付けてください。</span>
  <?php echo $PERM_HTML; ?>
  </div>
  <?php endif; ?>
  <?php if (isset($_GET['saved'])): ?><div class="stats">設定を保存しました（索引は次回検索時に自動で作り直されます）。</div><?php endif; ?>
  <?php if ($displayTrunc): ?><div class="warn">表示列は最大10個までです。先頭の10個を適用しています。</div><?php endif; ?>

  <?php if ($IDX_WARN): ?>
  <div class="warn" style="border-color:#e0245e;color:#b3003b;background:#fff0f3">警告：索引と設定の保存先が /Index になっていません（現在：<?php echo h($INDEX_DIR); ?>）。このままだとコンテナの更新・リセットで設定と索引が消えます。/Index フォルダへの書き込み権限（File Station で権限付与）か、ボリューム割り当てを確認してください。詳細は上部の「権限チェック」を見てください。</div>
  <?php endif; ?>
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
  <table class="rlist">
    <tr class="rhead">
      <?php if ($LAYOUT): ?>
        <?php foreach ($LAYOUT as $e): ?><th<?php echo in_array($e['key'], array('group','base_price','cost','spec'), true) ? ' class="hideS"' : ''; ?>><?php echo h(layout_label($e, $displayLabels)); ?></th><?php endforeach; ?>
      <?php elseif ($DISPLAY_COLS): ?>
        <?php foreach ($DISPLAY_COLS as $col): ?><th><?php echo h($displayLabels[$col]); ?><br><span class="loc"><?php echo h($col); ?>列</span></th><?php endforeach; ?>
      <?php else: ?>
        <th>プライス品番<br>タグ品番</th><th>品番コード</th><th>商品名</th>
        <th class="hideS">グループ</th><th>推定売価</th><th class="hideS">基準価格</th>
        <th class="hideS">推定原価</th><th class="hideS">スペック</th>
      <?php endif; ?>
      <?php if (!$LAYOUT): ?><th>Excel名</th><?php endif; ?>
      <?php if (!$LAYOUT && $IMG_ON): ?><th>画像</th><?php endif; ?>
    </tr>
    <?php foreach ($results as $r): ?>
    <tr>
      <?php if ($LAYOUT): ?>
        <?php foreach ($LAYOUT as $e): ?>
          <?php if ($e['key'] === 'image'): ?>
      <td data-label="画像"><?php echo img_cell($r, $SETTINGS['image_tpl'] ?? ''); ?></td>
          <?php elseif ($e['key'] === 'file'): ?>
      <td class="file" data-label="Excel名">
        <a href="?dl=<?php echo rawurlencode($r['file']); ?>"><?php echo h(basename($r['file'])); ?></a>
        <div class="loc"><?php echo (int)$r['row']; ?>行目</div>
        <details><summary>フォルダ・シート</summary>
        <div class="loc"><?php echo h(dirname($r['file']) === '.' ? '（親フォルダ）' : dirname($r['file'])); ?><br>
        <?php echo !empty($r['sno']) ? 'シート' . (int)$r['sno'] . '枚目' : h($r['sheet']); ?></div>
        </details>
      </td>
          <?php else: list($lv, $lc) = layout_value($r, $e); ?>
      <td data-label="<?php echo h(layout_label($e, $displayLabels)); ?>"<?php echo $lc !== '' ? ' class="' . $lc . '"' : ''; ?>><?php echo h($lv); ?><?php if ($e['key'] === 'name') echo age_div($r); ?></td>
          <?php endif; ?>
        <?php endforeach; ?>
      <?php elseif ($DISPLAY_COLS): ?>
        <?php foreach ($DISPLAY_COLS as $col): ?><td data-label="<?php echo h($displayLabels[$col]); ?>"><?php echo h(display_value($r, $col)); ?></td><?php endforeach; ?>
      <?php else: ?>
        <td class="price" data-label="プライス品番・タグ品番"><?php echo h($r['fields']['price']); ?></td>
        <td data-label="品番コード"><?php echo h($r['fields']['code']); ?></td>
        <td data-label="商品名"><?php echo h($r['fields']['name']); ?><?php echo age_div($r); ?></td>
        <td class="hideS" data-label="グループ"><?php echo h($r['fields']['group']); ?></td>
        <td class="num" data-label="推定売価"><?php echo h(fmt_money($r['fields']['sell'])); ?></td>
        <td class="num hideS" data-label="基準価格"><?php echo h(fmt_money($r['fields']['base_price'])); ?></td>
        <td class="num hideS" data-label="推定原価"><?php echo h(fmt_money($r['fields']['cost'])); ?></td>
        <td class="hideS" data-label="スペック"><?php echo h($r['fields']['spec']); ?></td>
      <?php endif; ?>
      <?php if (!$LAYOUT): ?>
      <td class="file" data-label="Excel名">
        <a href="?dl=<?php echo rawurlencode($r['file']); ?>"><?php echo h(basename($r['file'])); ?></a>
        <div class="loc"><?php echo (int)$r['row']; ?>行目</div>
        <details><summary>フォルダ・シート</summary>
        <div class="loc"><?php echo h(dirname($r['file']) === '.' ? '（親フォルダ）' : dirname($r['file'])); ?><br>
        <?php echo !empty($r['sno']) ? 'シート' . (int)$r['sno'] . '枚目' : h($r['sheet']); ?></div>
        </details>
      </td>
      <?php endif; ?>
      <?php if (!$LAYOUT && $IMG_ON): ?><td data-label="画像"><?php echo img_cell($r, $SETTINGS['image_tpl'] ?? ''); ?></td><?php endif; ?>
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
    <form method="post" action="" id="settingsForm">
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
        <label>表示レイアウト（JSON）</label>
        <textarea name="layout" rows="7" style="font-family:monospace;font-size:12.5px;padding:7px 9px;border:1px solid #bbb;border-radius:6px;width:100%;box-sizing:border-box"><?php echo h($SETTINGS['layout'] ?? ''); ?></textarea>
        <div class="note">結果表の列順と表示名をJSONで指定します（表示列より優先、最大15列）。key は price / code / name / group / sell / base_price / cost / spec / file、または {"col":"N","label":"好きな名前"} でシートの列を直接指定できます。空欄＝標準表示。<br>標準レイアウト：<code><?php echo h($LAYOUT_EXAMPLE); ?></code></div>
        <div class="note"><button type="button" class="edbtn" onclick="openLayoutEditor()">レイアウトを画面で編集（ポップアップ）</button> 並び替え・表示名の変更が画面でできます。</div>
        <label>画像URLテンプレート</label>
        <input type="text" name="image_tpl" value="<?php echo h($SETTINGS['image_tpl'] ?? ''); ?>" placeholder="例: https://hogehoge.com/gazou/{code}.JPG">
        <div class="note">結果に商品画像を表示します。{code}=品番コード、{price}=プライス品番、{name}=商品名、{group}=グループ に置換されます。未指定なら画像の枠自体を出しません。指定したのに画像が無い行は「ありません」と表示します。レイアウトの「画像」列でも使えます。</div>
        <label>ファイルの時期表示</label>
        <select name="show_age">
          <option value="0"<?php echo ($SETTINGS['show_age'] ?? '0') === '0' ? ' selected' : ''; ?>>表示しない</option>
          <option value="1"<?php echo ($SETTINGS['show_age'] ?? '0') === '1' ? ' selected' : ''; ?>>表示する</option>
        </select>
        <div class="note">商品名の下に、Excelファイルの更新日時から「2025年10月頃のファイル（約2ヶ月前）」のように時期を表示します。</div>
        <label>背景アニメーション</label>
        <select name="bg_mode">
          <option value="on"<?php echo $BG_MODE === 'on' ? ' selected' : ''; ?>>ON（標準の配色）</option>
          <option value="off"<?php echo $BG_MODE === 'off' ? ' selected' : ''; ?>>OFF</option>
          <option value="light"<?php echo $BG_MODE === 'light' ? ' selected' : ''; ?>>ライトテーマ</option>
          <option value="black"<?php echo $BG_MODE === 'black' ? ' selected' : ''; ?>>ブラックテーマ（画面も暗色・アニメは元の配色）</option>
        </select>
        <label>ノード数</label>
        <select name="bg_nodes">
          <option value="low"<?php echo $BG_NODES === 'low' ? ' selected' : ''; ?>>少ない（44）</option>
          <option value="mid"<?php echo $BG_NODES === 'mid' ? ' selected' : ''; ?>>中（88）</option>
          <option value="high"<?php echo $BG_NODES === 'high' ? ' selected' : ''; ?>>多い（160）</option>
        </select>
        <label>ノード表示</label>
        <select name="nodes_show">
          <option value="1"<?php echo ($SETTINGS['nodes_show'] ?? '1') === '1' ? ' selected' : ''; ?>>表示する</option>
          <option value="0"<?php echo ($SETTINGS['nodes_show'] ?? '1') === '0' ? ' selected' : ''; ?>>完全に非表示</option>
        </select>
        <div class="note">完全に非表示にすると、背景アニメーション自体を描画しません。</div>
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
    <div class="sbtn">
      <button type="button" class="sub2" onclick="openPerm()">権限チェック</button>
      <span class="note">フォルダの読み書き可否や /Index のマウント・書込状態をポップアップで確認できます。</span>
    </div>
  </details>
  <div style="text-align:right;color:#777;font-size:12.5px;margin-top:10px">
    Excel Multi-Sheet Search v<?php echo APP_VERSION; ?> ｜ <a href="#" onclick="openAbout();return false;">About</a>
  </div>
</div>
<div class="modal-bg" id="layoutModal">
  <div class="modal">
    <div style="display:flex;justify-content:space-between;align-items:center">
      <b>表示レイアウトの編集</b>
      <button type="button" class="edbtn" onclick="closeLayoutEditor()">閉じる</button>
    </div>
    <div style="color:#777;font-size:12px;margin:6px 0 10px">列の種類と表示名を設定し、↑↓で並び替えます。「決定」で保存して再読み込みします。</div>
    <div id="layoutRows"></div>
    <div style="margin-top:10px;display:flex;gap:8px;flex-wrap:wrap">
      <button type="button" class="edbtn" onclick="addLayoutRow()">＋ 列を追加</button>
      <button type="button" class="edbtn" onclick="resetLayout()">標準に戻す</button>
      <button type="button" class="edbtn" style="background:#1f6f43" onclick="applyLayout()">決定（保存して再読み込み）</button>
    </div>
  </div>
</div>
<script>
var LKEYS = <?php echo json_encode($KEY_LABELS, JSON_UNESCAPED_UNICODE | JSON_HEX_TAG | JSON_HEX_AMP); ?>;
var LSTATE = <?php echo json_encode($LAYOUT_EFFECTIVE, JSON_UNESCAPED_UNICODE | JSON_HEX_TAG | JSON_HEX_AMP); ?>;
function lesc(s){return String(s).replace(/&/g,'&amp;').replace(/"/g,'&quot;').replace(/</g,'&lt;');}
function renderLayoutRows(){
  var box = document.getElementById('layoutRows'); var html = '';
  LSTATE.forEach(function(e, i){
    var isCol = !e.key;
    html += '<div class="lrow">';
    html += '<select onchange="lsetType(' + i + ',this.value)">';
    Object.keys(LKEYS).forEach(function(k){
      html += '<option value="' + k + '"' + (e.key === k ? ' selected' : '') + '>' + lesc(LKEYS[k]) + '</option>';
    });
    html += '<option value="__col"' + (isCol ? ' selected' : '') + '>列を指定（A〜Z）</option>';
    html += '</select>';
    html += '<input type="text" style="width:52px" placeholder="列" value="' + lesc(e.col || '') + '"' + (isCol ? '' : ' disabled') + ' oninput="LSTATE[' + i + '].col=this.value">';
    html += '<input type="text" style="flex:1;min-width:130px" placeholder="表示名" value="' + lesc(e.label || '') + '" oninput="LSTATE[' + i + '].label=this.value">';
    html += '<button type="button" class="edbtn mv" onclick="lmove(' + i + ',-1)">↑</button>';
    html += '<button type="button" class="edbtn mv" onclick="lmove(' + i + ',1)">↓</button>';
    html += '<button type="button" class="edbtn mv" onclick="ldel(' + i + ')">✕</button>';
    html += '</div>';
  });
  box.innerHTML = html;
}
function lsetType(i, v){ if (v === '__col') { LSTATE[i].key = ''; } else { LSTATE[i].key = v; LSTATE[i].col = ''; } renderLayoutRows(); }
function lmove(i, d){ var j = i + d; if (j < 0 || j >= LSTATE.length) return; var t = LSTATE[i]; LSTATE[i] = LSTATE[j]; LSTATE[j] = t; renderLayoutRows(); }
function ldel(i){ LSTATE.splice(i, 1); renderLayoutRows(); }
function addLayoutRow(){ LSTATE.push({key:'name', col:'', label:''}); renderLayoutRows(); }
function openLayoutEditor(){ document.getElementById('layoutModal').style.display = 'block'; renderLayoutRows(); }
function closeLayoutEditor(){ document.getElementById('layoutModal').style.display = 'none'; }
function applyLayout(){
  var out = [];
  for (var i = 0; i < LSTATE.length; i++) {
    var e = LSTATE[i];
    if (e.key) { out.push({key: e.key, label: e.label || ''}); }
    else {
      var c = String(e.col || '').toUpperCase().trim();
      if (!/^[A-Z]{1,2}$/.test(c)) { alert('列の指定が正しくありません：' + (e.col || '（空）')); return; }
      out.push({col: c, label: e.label || ''});
    }
  }
  if (!out.length) { alert('列がありません'); return; }
  document.querySelector('textarea[name=layout]').value = JSON.stringify(out);
  document.getElementById('settingsForm').submit();
}
function resetLayout(){
  if (!confirm('レイアウトを標準に戻しますか？')) return;
  document.querySelector('textarea[name=layout]').value = '';
  document.getElementById('settingsForm').submit();
}
</script>
<div class="modal-bg" id="aboutModal">
  <div class="modal">
    <div style="display:flex;justify-content:space-between;align-items:center">
      <b>Excel Multi-Sheet Search</b>
      <button type="button" class="edbtn" onclick="closeAbout()">閉じる</button>
    </div>
    <p style="margin:10px 0">バージョン v<?php echo APP_VERSION; ?></p>
    <p>フォルダ内の Excel ファイル（.xlsx / .xlsm）を横断検索する Web ツールです。品番コードを入力すると、該当する行のプライス品番・商品名・価格などを一覧で表示します。</p>
    <ul style="margin:8px 0;padding-left:20px;line-height:1.7">
      <li>サブフォルダを含めて再帰検索（日本語フォルダ・ファイル名対応）</li>
      <li>索引（インデックス）方式で、2回目以降の検索が高速</li>
      <li>読み込むシート・検索列・表示列・表示レイアウト（列順と表示名）を設定可能</li>
      <li>Docker イメージで配布（Synology NAS / Linux で動作）</li>
    </ul>
    <p style="line-height:1.8">
      GitHub：<a href="https://github.com/Yakushi894/Excel_Multi-Sheet_Search" target="_blank" rel="noopener">Yakushi894/Excel_Multi-Sheet_Search</a><br>
      Docker Hub：<a href="https://hub.docker.com/r/yuichi894/excel-multi-sheet-search" target="_blank" rel="noopener">yuichi894/excel-multi-sheet-search</a>
    </p>
    <p style="color:#777;font-size:12.5px">License: MIT © Yakushi894</p>
  </div>
</div>
<script>
function openAbout(){ document.getElementById('aboutModal').style.display = 'block'; }
function closeAbout(){ document.getElementById('aboutModal').style.display = 'none'; }
</script>
<div class="modal-bg" id="permModal" onclick="if(event.target===this)closePerm()">
  <div class="modal">
    <?php echo $PERM_HTML; ?>
    <div style="margin-top:12px;text-align:right"><button type="button" class="edbtn" onclick="closePerm()">閉じる</button></div>
  </div>
</div>
<div class="modal-bg" id="imgModal" onclick="if(event.target.id!=='imgPopupImg')closeImgPopup()">
  <div style="text-align:center;padding:4vh 10px">
    <img id="imgPopupImg" src="" alt="" style="max-width:96vw;max-height:86vh;border-radius:8px;box-shadow:0 8px 30px rgba(0,0,0,.4)">
    <div style="margin-top:10px"><button type="button" class="edbtn" onclick="closeImgPopup()">閉じる</button></div>
  </div>
</div>
<script>
function openPerm(){ document.getElementById('permModal').style.display = 'block'; }
function closePerm(){ document.getElementById('permModal').style.display = 'none'; }
function openImgPopup(src){ document.getElementById('imgPopupImg').src = src; document.getElementById('imgModal').style.display = 'block'; }
function closeImgPopup(){ document.getElementById('imgModal').style.display = 'none'; document.getElementById('imgPopupImg').src = ''; }
</script>
</body>
</html>
