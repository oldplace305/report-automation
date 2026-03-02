#!/usr/bin/env python3
"""
居宅療養管理指導 報告書PDF自動出力ツール

Solamichiから介護報告書PDFを自動でダウンロード・リネーム・仕分けするCLIツール。
"""

import argparse
import calendar
import glob
import json
import logging
import os
import re
import shutil
import subprocess
import sys
import time
import traceback
import unicodedata
from datetime import datetime

from selenium import webdriver
from selenium.common.exceptions import (
    NoAlertPresentException,
    NoSuchWindowException,
    TimeoutException,
    UnexpectedAlertPresentException,
    WebDriverException,
)
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait

# =============================================================================
# 定数
# =============================================================================
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
CONFIG_PATH = os.path.join(BASE_DIR, "config.json")
PROGRESS_PATH = os.path.join(BASE_DIR, "progress.json")
LOG_PATH = os.path.join(BASE_DIR, "log.txt")
ERRORS_DIR = os.path.join(BASE_DIR, "errors")


# =============================================================================
# ログ設定
# =============================================================================
def setup_logging():
    """ファイル + コンソールの二重出力ログを設定する。"""
    log_logger = logging.getLogger("report_tool")
    log_logger.setLevel(logging.DEBUG)

    formatter = logging.Formatter(
        "[%(asctime)s] [%(levelname)-5s] %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S",
    )

    # ファイルハンドラ（追記）
    fh = logging.FileHandler(LOG_PATH, encoding="utf-8")
    fh.setLevel(logging.DEBUG)
    fh.setFormatter(formatter)
    log_logger.addHandler(fh)

    # コンソールハンドラ
    ch = logging.StreamHandler(sys.stdout)
    ch.setLevel(logging.INFO)
    ch.setFormatter(formatter)
    log_logger.addHandler(ch)

    return log_logger


logger = setup_logging()


# =============================================================================
# 設定読み込み
# =============================================================================
def load_config():
    """config.json を読み込み、チルダ展開を適用して返す。"""
    with open(CONFIG_PATH, "r", encoding="utf-8") as f:
        config = json.load(f)

    # チルダ展開
    for key in ["download_folder", "output_base_folder", "chrome_user_data_dir", "chrome_debug_user_data_dir"]:
        if key in config:
            config[key] = os.path.expanduser(config[key])

    return config


# =============================================================================
# 進捗管理
# =============================================================================
def load_progress():
    """progress.json を読み込む。存在しなければ None を返す。"""
    if not os.path.exists(PROGRESS_PATH):
        return None
    with open(PROGRESS_PATH, "r", encoding="utf-8") as f:
        return json.load(f)


def save_progress(progress_data):
    """progress.json に書き込む。"""
    progress_data["last_run"] = datetime.now().isoformat()
    with open(PROGRESS_PATH, "w", encoding="utf-8") as f:
        json.dump(progress_data, f, ensure_ascii=False, indent=2)


def reset_progress():
    """progress.json をリセットする。"""
    if os.path.exists(PROGRESS_PATH):
        os.remove(PROGRESS_PATH)


# =============================================================================
# ユーティリティ
# =============================================================================
def normalize_name(name):
    """macOSのNFD正規化に対応するため、NFC正規化して返す。"""
    return unicodedata.normalize("NFC", name.strip())


def get_target_month_label(year, month):
    """例: '2026年1月分'"""
    return f"{year}年{month}月分"


def get_date_range(year, month):
    """対象月の1日と末日を返す。(YYYY-MM-DD 形式 — HTML input[type=date] に合わせる)"""
    last_day = calendar.monthrange(year, month)[1]
    from_date = f"{year}-{month:02d}-01"
    to_date = f"{year}-{month:02d}-{last_day:02d}"
    return from_date, to_date


def make_output_dir(config, year, month):
    """出力先ディレクトリを作成して返す。"""
    label = get_target_month_label(year, month)
    output_dir = os.path.join(config["output_base_folder"], label)
    os.makedirs(output_dir, exist_ok=True)
    return output_dir


def generate_unique_filename(output_dir, care_manager_name, year, month):
    """
    同姓のケアマネがいる場合に連番を付与したファイルパスを返す。
    例: 浅田様_2026年1月分.pdf, 浅田様(2)_2026年1月分.pdf
    """
    label = get_target_month_label(year, month)
    base_name = f"{care_manager_name}様_{label}.pdf"
    full_path = os.path.join(output_dir, base_name)

    if not os.path.exists(full_path):
        return full_path

    # 連番付与
    counter = 2
    while True:
        numbered_name = f"{care_manager_name}様({counter})_{label}.pdf"
        full_path = os.path.join(output_dir, numbered_name)
        if not os.path.exists(full_path):
            return full_path
        counter += 1


def check_chrome_debuggable(port=9222):
    """リモートデバッグ付きChromeに接続できるかチェックする。"""
    import urllib.request
    try:
        urllib.request.urlopen(f"http://localhost:{port}/json/version", timeout=3)
        return True
    except Exception:
        return False


def wait_for_download(download_dir, timeout=60, check_interval=2):
    """
    ダウンロード完了を待機する。

    1. .crdownload ファイルが存在しなくなるまで待つ
    2. 直近30秒以内に作成された .pdf ファイルを取得
    """
    end_time = time.time() + timeout

    while time.time() < end_time:
        # .crdownload がまだあるならダウンロード中
        crdownload_files = glob.glob(os.path.join(download_dir, "*.crdownload"))
        if not crdownload_files:
            # ダウンロード完了 → 最新のPDFを取得
            pdf_files = glob.glob(os.path.join(download_dir, "*.pdf"))
            if pdf_files:
                latest = max(pdf_files, key=os.path.getctime)
                # 直近30秒以内に作成されたファイルのみ対象
                if time.time() - os.path.getctime(latest) < 30:
                    return latest
        time.sleep(check_interval)

    raise TimeoutError("ダウンロードがタイムアウトしました")


def save_error_screenshot(driver, care_manager_name):
    """エラー時のスクリーンショットを保存する。"""
    os.makedirs(ERRORS_DIR, exist_ok=True)
    timestamp = datetime.now().strftime("%Y%m%d%H%M%S")
    filename = f"error_{care_manager_name}_{timestamp}.png"
    filepath = os.path.join(ERRORS_DIR, filename)
    try:
        driver.save_screenshot(filepath)
        logger.info(f"スクリーンショット保存: {filepath}")
    except Exception as e:
        logger.error(f"スクリーンショット保存失敗: {e}")


# =============================================================================
# Selenium操作
# =============================================================================
def create_driver(config):
    """既にリモートデバッグモードで起動しているChromeに接続する。"""
    options = Options()
    options.debugger_address = f"localhost:{config.get('chrome_debug_port', 9222)}"

    driver = webdriver.Chrome(options=options)
    driver.implicitly_wait(5)
    return driver


def navigate_to_page(driver, config):
    """Solamichiの介護報告書一括印刷画面へ遷移する。"""
    driver.get(config["solamichi_url"])
    time.sleep(config["wait_time"]["page_load"])
    logger.info(f"ページ遷移完了: {config['solamichi_url']}")


def set_search_conditions(driver, config, from_date, to_date, wait_timeout):
    """
    検索条件を設定する。
    - 作成日(From/To)を入力
    - 宛先「ケアマネ」ラジオボタンを選択
    """
    wait = WebDriverWait(driver, wait_timeout)

    # 作成日(From/To) — type="date" の input 要素
    date_inputs = wait.until(
        EC.presence_of_all_elements_located(
            (By.XPATH, '//input[@type="date"]')
        )
    )
    if len(date_inputs) < 2:
        raise RuntimeError("日付入力フィールドが見つかりません")

    # JavaScriptで値を設定（type="date"はsend_keysが不安定なため）
    # Vue.jsではinput + change両方のイベントを発火させる必要がある
    set_date_js = """
        var el = arguments[0];
        var nativeInputValueSetter = Object.getOwnPropertyDescriptor(
            window.HTMLInputElement.prototype, 'value'
        ).set;
        nativeInputValueSetter.call(el, arguments[1]);
        el.dispatchEvent(new Event('input', { bubbles: true }));
        el.dispatchEvent(new Event('change', { bubbles: true }));
    """
    driver.execute_script(set_date_js, date_inputs[0], from_date)
    driver.execute_script(set_date_js, date_inputs[1], to_date)

    # 宛先「ケアマネ」ラジオボタン — Vue.jsの場合、inputを直接クリックすると
    # リアクティブ更新が発火しないため、labelをクリックする必要がある
    driver.execute_script(
        'document.querySelector(\'label[for="destination_care_manager"]\').click()'
    )

    # ケアマネモードのUI表示切替を待機（「ケアマネ：」ラベルが表示されるまで）
    wait.until(
        lambda d: d.execute_script('''
            var spans = document.querySelectorAll('.input-title');
            for (var s of spans) {
                if (s.textContent.trim() === 'ケアマネ：') return true;
            }
            return false;
        ''')
    )

    time.sleep(0.5)
    logger.debug("検索条件設定完了")


def open_care_manager_modal(driver, config, wait_timeout):
    """「ケアマネ：選択」ボタンをクリックしてモーダルを開く。"""
    # ケアマネモードでは「施設：選択」「ケアマネ：選択」が表示される。
    # 「ケアマネ：」ラベルの隣にある「選択」ボタンを特定してクリックする。
    clicked = driver.execute_script('''
        var spans = document.querySelectorAll('.input-title');
        for (var s of spans) {
            if (s.textContent.trim() === 'ケアマネ：') {
                var container = s.parentElement;
                var btn = container.querySelector('button');
                if (btn) { btn.click(); return true; }
            }
        }
        return false;
    ''')
    if not clicked:
        logger.warning("ケアマネ選択ボタンが見つかりません。フォールバック方法を試行...")
        # フォールバック: 2番目の「選択」ボタンをクリック
        driver.execute_script('''
            var btns = document.querySelectorAll('button.btn');
            var selectBtns = [];
            for (var b of btns) {
                if (b.textContent.trim() === '選択' && b.offsetParent !== null) {
                    selectBtns.push(b);
                }
            }
            if (selectBtns.length >= 2) selectBtns[1].click();
            else if (selectBtns.length === 1) selectBtns[0].click();
        ''')
    time.sleep(config["wait_time"]["page_load"])

    # モーダルが開くのを待つ
    WebDriverWait(driver, wait_timeout).until(
        EC.presence_of_element_located(
            (By.CSS_SELECTOR, ".modal-content .edit-area-input")
        )
    )


def get_care_manager_list(driver, config, wait_timeout):
    """
    「選択」ボタンからモーダルを開き、ケアマネリストを取得する。
    ＝＝＝＝＝◯◯＝＝＝＝＝ 形式の見出し行はスキップする。

    戻り値:
        list of dict: [{"name": "浦田", "index": 0}, {"name": "山下", "index": 19}, ...]
        index は .edit-area-input 要素の全体通し番号（重複名を区別するために使用）
    """
    open_care_manager_modal(driver, config, wait_timeout)

    # モーダル内の input.edit-area-input の value からケアマネ名を取得
    names = driver.execute_script('''
        var inputs = document.querySelectorAll('.modal-content .edit-area-input');
        var result = [];
        for (var i = 0; i < inputs.length; i++) {
            result.push({value: inputs[i].value, index: i});
        }
        return result;
    ''')

    # フィルタリング
    heading_pattern = re.compile(r"^[＝=]+.*[＝=]+$")
    care_managers = []

    for item in names:
        name = normalize_name(item["value"])
        if not name:
            continue
        # 見出し行（＝＝＝＝＝◯◯＝＝＝＝＝）をスキップ
        if heading_pattern.match(name):
            continue
        # 注意書きや長すぎるテキストを除外
        if len(name) > 15 or "※" in name:
            continue
        care_managers.append({"name": name, "index": item["index"]})

    # モーダルを閉じる
    close_modal(driver)
    time.sleep(1)

    logger.info(f"ケアマネリスト取得完了: 全{len(care_managers)}名")
    return care_managers


def close_modal(driver):
    """モーダルを閉じる。modal-maskクリック or ESCキー。"""
    # 方法1: モーダルマスク（背景）をクリック
    try:
        mask = driver.find_element(By.CSS_SELECTOR, ".modal-mask")
        if mask.is_displayed():
            driver.execute_script("arguments[0].click();", mask)
            return
    except Exception:
        pass

    # 方法2: ESCキー
    try:
        driver.find_element(By.TAG_NAME, "body").send_keys(Keys.ESCAPE)
    except Exception:
        pass


def set_care_manager_name(driver, cm_name):
    """
    ケアマネ名をテキスト入力欄に直接設定する。

    モーダルでの行選択は「ケアマネリスト管理」用であり、
    検索条件としてケアマネ名を設定するには入力欄に直接値をセットする。
    """
    result = driver.execute_script('''
        var spans = document.querySelectorAll('.input-title');
        for (var s of spans) {
            if (s.textContent.trim() === 'ケアマネ：') {
                var container = s.parentElement;
                var input = container.querySelector('input');
                if (input) {
                    var nativeInputValueSetter = Object.getOwnPropertyDescriptor(
                        window.HTMLInputElement.prototype, 'value'
                    ).set;
                    nativeInputValueSetter.call(input, arguments[0]);
                    input.dispatchEvent(new Event('input', { bubbles: true }));
                    input.dispatchEvent(new Event('change', { bubbles: true }));
                    return true;
                }
            }
        }
        return false;
    ''', cm_name)

    if not result:
        raise RuntimeError(f"ケアマネ入力欄が見つかりません")

    logger.debug(f"ケアマネ入力欄に '{cm_name}' を設定")


def select_care_manager_and_search(driver, config, care_manager_info, wait_timeout):
    """
    特定のケアマネを選択し、検索を実行する。

    ケアマネ名を入力欄に直接設定してから「検索」ボタンをクリックする。

    戻り値:
        "found"     - 検索結果あり
        "not_found" - 報告書なし（アラート）
    """
    cm_name = care_manager_info["name"]

    # ケアマネ名を入力欄に設定
    set_care_manager_name(driver, cm_name)
    time.sleep(0.5)

    # 「検索」ボタンをクリック
    driver.execute_script("""
        var btns = document.querySelectorAll('button');
        for (var b of btns) {
            if (b.textContent.trim() === '\u691c\u7d22' && b.offsetParent !== null) {
                b.click();
                return;
            }
        }
    """)
    time.sleep(config["wait_time"]["page_load"])
    logger.debug(f"「検索」クリック完了（ケアマネ: {cm_name}）")

    # 分岐判定: アラート（報告書なし）or 検索結果あり
    try:
        WebDriverWait(driver, 3).until(EC.alert_is_present())
        alert = driver.switch_to.alert
        alert_text = alert.text
        logger.debug(f"アラート検出: {alert_text}")
        alert.accept()
        return "not_found"
    except TimeoutException:
        # アラートなし = 検索結果あり
        return "found"


def click_select_all(driver, wait_timeout):
    """「すべて選択」ボタンをクリックして全報告書を選択する。"""
    wait = WebDriverWait(driver, wait_timeout)
    select_all_btn = wait.until(
        EC.element_to_be_clickable(
            (By.XPATH, '//button[contains(text(), "すべて選択")]')
        )
    )
    driver.execute_script("arguments[0].click();", select_all_btn)
    time.sleep(1)
    logger.debug("「すべて選択」クリック完了")


def click_print_button(driver, wait_timeout):
    """「白紙を含まずに印刷」ボタンをクリックする。"""
    wait = WebDriverWait(driver, wait_timeout)
    print_btn = wait.until(
        EC.element_to_be_clickable(
            (By.XPATH, '//button[contains(text(), "白紙を含まずに印刷")]')
        )
    )
    driver.execute_script("arguments[0].click();", print_btn)


def click_download_button_cdp(driver, care_manager_name=""):
    """
    印刷履歴テーブルの対象行のダウンロードボタンをCDP経由でクリックする。

    care_manager_name が指定されている場合、印刷条件にケアマネ名が含まれる行を
    特定してそのボタンをクリックする。見つからない場合は最新行をフォールバック。

    Nuxt3 (Vue3) のイベントリスナーは execute_script の element.click() では
    発火しないことがあるため、CDP の Input.dispatchMouseEvent を使って
    実際のマウスクリックをシミュレートする。
    """
    # ケアマネ名で正しい行のボタンを特定し、座標を取得
    coords = driver.execute_script("""
        var tbody = document.querySelector('.table_bg tbody');
        if (!tbody) return null;
        var rows = tbody.querySelectorAll('tr');
        if (rows.length === 0) return null;

        var cmName = arguments[0];
        var targetRow = null;

        // ケアマネ名で行を特定
        if (cmName) {
            for (var r of rows) {
                var cells = r.querySelectorAll('td');
                if (cells.length >= 5 && cells[4].textContent.indexOf(cmName) >= 0) {
                    // ステータスが「処理完了」の行のみ対象
                    if (cells[1].textContent.trim() === '処理完了') {
                        targetRow = r;
                        break;
                    }
                }
            }
        }

        // フォールバック: 最新行
        if (!targetRow) targetRow = rows[0];

        var btn = targetRow.querySelector('button');
        if (!btn) return null;
        btn.scrollIntoView({block: 'center'});
        var rect = btn.getBoundingClientRect();
        return {
            x: Math.round(rect.left + rect.width / 2),
            y: Math.round(rect.top + rect.height / 2)
        };
    """, care_manager_name)

    if not coords:
        raise RuntimeError("ダウンロードボタンの座標が取得できません")

    x = coords["x"]
    y = coords["y"]

    # CDP経由でマウスクリックを送信
    driver.execute_cdp_cmd("Input.dispatchMouseEvent", {
        "type": "mousePressed",
        "x": x,
        "y": y,
        "button": "left",
        "clickCount": 1,
    })
    driver.execute_cdp_cmd("Input.dispatchMouseEvent", {
        "type": "mouseReleased",
        "x": x,
        "y": y,
        "button": "left",
        "clickCount": 1,
    })
    logger.debug(f"ダウンロードボタン: CDPクリック実行 ({x}, {y})")


def wait_for_print_history_and_download(driver, config, care_manager_name=""):
    """
    「白紙を含まずに印刷」クリック後の処理:

    1. 印刷履歴ページへの遷移を待機（自動遷移 or 手動遷移）
    2. 最新の印刷ジョブ（ケアマネ名で識別）のステータスが「処理完了」になるまでポーリング
    3. ダウンロードボタンをクリックしてPDFをダウンロード
    """
    pdf_timeout = config["wait_time"]["pdf_generation_timeout"]
    page_load_wait = config["wait_time"]["page_load"]

    # --- Step 1: 印刷履歴ページへの遷移を待機 ---
    logger.debug("印刷履歴ページへの遷移を待機中...")
    end_time = time.time() + 30
    while time.time() < end_time:
        current_url = driver.current_url
        if "print-history" in current_url:
            logger.debug("印刷履歴ページに遷移しました")
            break
        # 印刷履歴タブが表示されていればクリックして遷移
        try:
            switched = driver.execute_script("""
                var links = document.querySelectorAll('.tab_link');
                for (var a of links) {
                    if (a.textContent.trim().indexOf('\u5370\u5237\u5c65\u6b74') >= 0) {
                        a.click();
                        return true;
                    }
                }
                return false;
            """)
            if switched:
                time.sleep(page_load_wait)
                if "print-history" in driver.current_url:
                    break
        except Exception:
            pass
        time.sleep(2)
    else:
        raise RuntimeError("印刷履歴ページへ遷移できませんでした")

    time.sleep(page_load_wait)

    # --- Step 2: 自分のジョブのステータスが「処理完了」になるまで待機 ---
    # 印刷条件の列にケアマネ名が含まれる最新の行を自分のジョブとして識別する
    logger.debug(f"PDF生成完了を待機中...（ケアマネ: {care_manager_name}）")
    poll_start = time.time()
    end_time = poll_start + pdf_timeout
    check_count = 0
    CHECKS_BEFORE_RELOAD = 15  # 2秒×15回=30秒ごとにリロード

    while time.time() < end_time:
        # 全行を走査して、ケアマネ名が含まれる最新行のステータスを取得
        status = driver.execute_script("""
            var table = document.querySelector('.table_bg');
            if (!table) return {error: 'no table'};
            var tbody = table.querySelector('tbody');
            if (!tbody) return {error: 'no tbody'};
            var rows = tbody.querySelectorAll('tr');
            if (rows.length === 0) return {error: 'no rows'};

            var cmName = arguments[0];

            // ケアマネ名で自分のジョブを探す（最新の行から順に）
            for (var i = 0; i < rows.length; i++) {
                var cells = rows[i].querySelectorAll('td');
                if (cells.length < 6) continue;

                var condition = cells[4].textContent.trim();  // 印刷条件の列

                // ケアマネ名が印刷条件に含まれているか確認
                if (cmName && condition.indexOf(cmName) >= 0) {
                    var status = cells[1].textContent.trim();
                    var hasBtn = rows[i].querySelector('button') !== null;
                    return {
                        status: status,
                        hasButton: hasBtn,
                        rowIndex: i,
                        condition: condition.substring(0, 80),
                        matched: true
                    };
                }
            }

            // ケアマネ名で見つからない場合は最新行を使用（フォールバック）
            var cells = rows[0].querySelectorAll('td');
            if (cells.length < 6) return {error: 'insufficient cells'};
            return {
                status: cells[1].textContent.trim(),
                hasButton: rows[0].querySelector('button') !== null,
                rowIndex: 0,
                condition: cells[4].textContent.trim().substring(0, 80),
                matched: false
            };
        """, care_manager_name)

        if isinstance(status, dict) and status.get("status") == "処理完了":
            row_idx = status.get("rowIndex", 0)
            matched = status.get("matched", False)
            logger.debug(
                f"PDF生成完了（ステータス: 処理完了, 行: {row_idx}, "
                f"ケアマネ一致: {matched}）"
            )
            break
        elif isinstance(status, dict) and "error" in status:
            logger.debug(f"テーブル待機中: {status['error']}")
        else:
            current_status = status.get('status', '不明') if isinstance(status, dict) else status
            elapsed_sec = int(time.time() - poll_start)
            matched = status.get("matched", False) if isinstance(status, dict) else False
            logger.debug(
                f"ステータス待機中: {current_status}（経過{elapsed_sec}秒, "
                f"ケアマネ一致: {matched}）"
            )

        check_count += 1

        if check_count % CHECKS_BEFORE_RELOAD == 0:
            logger.debug(f"ステータス未変化のためページをリロード（{check_count}回チェック済み）")
            driver.refresh()
            time.sleep(page_load_wait)
        else:
            time.sleep(2)
    else:
        raise TimeoutError(f"PDF生成が{pdf_timeout}秒以内に完了しませんでした")

    # --- Step 3: ダウンロードボタンをクリック ---
    logger.debug("ダウンロードボタンをクリック...")
    time.sleep(1)

    original_window = driver.current_window_handle
    click_download_button_cdp(driver, care_manager_name)

    # --- Step 4: PDFプレビュータブからダウンロード ---
    time.sleep(page_load_wait + 1)
    download_pdf_from_preview_tab(driver, config, original_window)


def download_pdf_from_preview_tab(driver, config, original_window):
    """
    PDFプレビュータブ (/pdf-preview) からPDFをダウンロードする。

    Solamichiのダウンロードボタンは新しいタブに /pdf-preview ページを開き、
    <object type="application/pdf" data="blob:..."> でPDFを表示する。
    この blob URL を fetch → createObjectURL → <a download> クリック で
    ダウンロードフォルダにPDFを保存する。
    """
    page_load_wait = config["wait_time"]["page_load"]

    # 新しいタブを探す
    new_tab = None
    for w in driver.window_handles:
        if w != original_window:
            new_tab = w
            break

    if not new_tab:
        logger.warning("PDFプレビュータブが開かれませんでした（直接ダウンロードの可能性）")
        return

    driver.switch_to.window(new_tab)
    time.sleep(page_load_wait)

    # blob URL を取得（最大30秒待機）
    blob_url = None
    end_time = time.time() + 30
    while time.time() < end_time:
        blob_url = driver.execute_script("""
            var obj = document.querySelector('object[type="application/pdf"]');
            return obj ? obj.data : null;
        """)
        if blob_url and blob_url.startswith("blob:"):
            break
        time.sleep(1)

    if not blob_url:
        logger.warning("PDFプレビューのblob URLが取得できませんでした")
        driver.close()
        driver.switch_to.window(original_window)
        return

    logger.debug(f"blob URL取得: {blob_url[:80]}...")

    # blob URL を fetch してダウンロード
    driver.execute_script("""
        var blobUrl = arguments[0];
        fetch(blobUrl)
            .then(function(response) { return response.blob(); })
            .then(function(blob) {
                var a = document.createElement('a');
                a.href = URL.createObjectURL(blob);
                a.download = 'report.pdf';
                document.body.appendChild(a);
                a.click();
                document.body.removeChild(a);
                URL.revokeObjectURL(a.href);
            });
    """, blob_url)
    logger.debug("blob fetch → ダウンロード実行")

    time.sleep(3)

    # プレビュータブを閉じる
    try:
        driver.close()
    except Exception:
        pass  # タブが既に閉じている場合は無視
    driver.switch_to.window(original_window)


def navigate_back_to_main(driver, config):
    """介護報告書一括印刷画面に戻る。"""
    driver.get(config["solamichi_url"])
    time.sleep(config["wait_time"]["page_load"])


# =============================================================================
# メイン処理
# =============================================================================
def process_care_manager(
    driver, config, care_manager_info, index, total, from_date, to_date, output_dir
):
    """
    1人のケアマネに対する一連の処理を行う。

    care_manager_info: dict {"name": "...", "index": N}

    戻り値:
        ("success", "")           - PDF保存成功
        ("skipped", "報告書なし")  - 報告書なし
        ("failed", "理由")        - エラー発生
    """
    year = config["target_year"]
    month = config["target_month"]
    wait_timeout = config["wait_time"]["element_timeout"]
    care_manager_name = care_manager_info["name"]

    logger.info(f"[{index}/{total}] {care_manager_name}様 の処理開始")

    try:
        # 1. メインページへ移動（毎回リセットするため）
        navigate_back_to_main(driver, config)

        # 2. 検索条件を設定
        set_search_conditions(driver, config, from_date, to_date, wait_timeout)

        # 3. ケアマネを選択して検索
        result = select_care_manager_and_search(
            driver, config, care_manager_info, wait_timeout
        )

        if result == "not_found":
            logger.info(
                f"[{index}/{total}] {care_manager_name}様: 報告書なし（スキップ）"
            )
            return ("skipped", "報告書なし")

        # 4. 「すべて選択」→「白紙を含まずに印刷」
        click_select_all(driver, wait_timeout)
        click_print_button(driver, wait_timeout)
        time.sleep(config["wait_time"]["page_load"])

        # 5. 印刷履歴ページで処理完了を待ち、ダウンロード
        wait_for_print_history_and_download(driver, config, care_manager_name)
        time.sleep(config["wait_time"]["page_load"])

        # 6. ダウンロード完了を待機
        downloaded_file = wait_for_download(
            config["download_folder"],
            timeout=config["wait_time"]["download_timeout"],
            check_interval=config["wait_time"]["download_check_interval"],
        )

        # 7. PDFをリネーム・移動
        dest_path = generate_unique_filename(
            output_dir, care_manager_name, year, month
        )
        shutil.move(downloaded_file, dest_path)

        dest_filename = os.path.basename(dest_path)
        logger.info(
            f"[{index}/{total}] ✓ {care_manager_name}様: "
            f"PDF保存完了 → {dest_filename}"
        )
        return ("success", "")

    except TimeoutError as e:
        reason = f"タイムアウト（{e}）"
        logger.error(f"[{index}/{total}] ✗ {care_manager_name}様: {reason}")
        save_error_screenshot(driver, care_manager_name)
        return ("failed", reason)

    except TimeoutException as e:
        reason = f"タイムアウト（要素待機）"
        logger.error(f"[{index}/{total}] ✗ {care_manager_name}様: {reason}")
        save_error_screenshot(driver, care_manager_name)
        return ("failed", reason)

    except RuntimeError as e:
        reason = str(e)
        logger.error(f"[{index}/{total}] ✗ {care_manager_name}様: {reason}")
        save_error_screenshot(driver, care_manager_name)
        return ("failed", reason)

    except Exception as e:
        reason = f"予期せぬエラー: {e}"
        logger.error(f"[{index}/{total}] ✗ {care_manager_name}様: {reason}")
        logger.debug(traceback.format_exc())
        save_error_screenshot(driver, care_manager_name)
        return ("failed", reason)

    finally:
        # アラートが残っていたら閉じる
        try:
            alert = driver.switch_to.alert
            alert.accept()
        except NoAlertPresentException:
            pass

        # メインウィンドウに戻る（タブが複数開いている場合の安全策）
        try:
            if len(driver.window_handles) > 1:
                main_window = driver.window_handles[0]
                for handle in driver.window_handles[1:]:
                    driver.switch_to.window(handle)
                    driver.close()
                driver.switch_to.window(main_window)
        except Exception:
            pass


def run_dry_run(driver, config):
    """ドライランモード: ケアマネリストの取得と件数確認のみ行う。"""
    year = config["target_year"]
    month = config["target_month"]
    from_date, to_date = get_date_range(year, month)
    wait_timeout = config["wait_time"]["element_timeout"]

    logger.info("[DRY RUN] ドライランモードで実行中...")

    # 検索条件設定
    set_search_conditions(driver, config, from_date, to_date, wait_timeout)

    # ケアマネリスト取得
    care_managers = get_care_manager_list(driver, config, wait_timeout)

    print()
    print(f"[DRY RUN] ケアマネリスト取得完了: 全{len(care_managers)}名")
    for i, cm in enumerate(care_managers, 1):
        print(f"[DRY RUN]  {i:>2}. {cm['name']}")
    print()
    print("[DRY RUN] 実行時は --dry-run を外してください")


def run_main(driver, config):
    """通常実行モード: 全ケアマネのPDFを処理する。"""
    year = config["target_year"]
    month = config["target_month"]
    month_label = get_target_month_label(year, month)
    from_date, to_date = get_date_range(year, month)
    wait_timeout = config["wait_time"]["element_timeout"]

    start_time = time.time()

    # 出力ディレクトリ作成
    output_dir = make_output_dir(config, year, month)

    # 検索条件設定
    set_search_conditions(driver, config, from_date, to_date, wait_timeout)

    # ケアマネリスト取得
    care_managers = get_care_manager_list(driver, config, wait_timeout)

    if not care_managers:
        logger.warning("ケアマネリストが空です。処理を終了します。")
        return

    # ケアマネごとの一意キー（同姓を区別）: "名前_index"
    def cm_key(cm_info):
        return f"{cm_info['name']}_{cm_info['index']}"

    # 進捗確認（前回の続きから再開するか）
    progress = load_progress()
    completed_set = set()
    skipped_set = set()
    failed_dict = {}

    if progress and progress.get("target_month") == month_label:
        completed_set = set(progress.get("completed", []))
        skipped_set = set(progress.get("skipped", []))
        remaining = [
            cm for cm in care_managers
            if cm_key(cm) not in completed_set and cm_key(cm) not in skipped_set
        ]
        if remaining and len(completed_set) > 0:
            answer = input(
                f"\n前回の続き ({len(completed_set)}件処理済み) "
                f"から再開しますか？ (y/n): "
            ).strip()
            if answer.lower() == "y":
                logger.info(
                    f"前回の続きから再開: {len(completed_set)}件処理済み, "
                    f"残り{len(remaining)}件"
                )
            else:
                logger.info("最初からやり直します")
                completed_set = set()
                skipped_set = set()
                failed_dict = {}
                reset_progress()
        else:
            reset_progress()
    else:
        reset_progress()

    total = len(care_managers)
    success_count = len(completed_set)
    skip_count = len(skipped_set)
    fail_count = 0

    logger.info(f"処理開始: 全{total}名")

    for i, cm_info in enumerate(care_managers, 1):
        key = cm_key(cm_info)

        # 既に処理済み or スキップ済みならスキップ
        if key in completed_set or key in skipped_set:
            continue

        result_status, result_reason = process_care_manager(
            driver, config, cm_info, i, total, from_date, to_date, output_dir
        )

        if result_status == "success":
            completed_set.add(key)
            success_count += 1
        elif result_status == "skipped":
            skipped_set.add(key)
            skip_count += 1
        elif result_status == "failed":
            failed_dict[cm_info["name"]] = result_reason
            fail_count += 1

        # 進捗を保存（毎回）
        save_progress(
            {
                "target_month": month_label,
                "all_care_managers": [cm_key(c) for c in care_managers],
                "completed": list(completed_set),
                "skipped": list(skipped_set),
                "failed": list(failed_dict.keys()),
                "last_index": i,
            }
        )

    # 処理完了サマリー
    elapsed = time.time() - start_time
    elapsed_min = int(elapsed // 60)
    elapsed_sec = int(elapsed % 60)

    summary = f"""
========================================
  処理完了サマリー
========================================
対象月:   {month_label}
処理時間: {elapsed_min}分{elapsed_sec:02d}秒
総数:     {total}名

✓ 成功:    {success_count}件
  スキップ: {skip_count}件（報告書なし）
✗ 失敗:    {fail_count}件
"""

    if failed_dict:
        summary += "\n失敗リスト:\n"
        for name, reason in failed_dict.items():
            summary += f"  - {name}様: {reason}\n"

    summary += f"\n出力先: {output_dir}\n"
    summary += "========================================"

    logger.info("=== 処理完了 ===")
    logger.info(
        f"成功: {success_count}件 / スキップ: {skip_count}件 / 失敗: {fail_count}件"
    )

    print(summary)


# =============================================================================
# エントリポイント
# =============================================================================
def main():
    parser = argparse.ArgumentParser(
        description="居宅療養管理指導 報告書PDF自動出力ツール"
    )
    parser.add_argument(
        "--dry-run",
        action="store_true",
        help="ケアマネリストの取得と件数確認のみ行う（ダウンロードは行わない）",
    )
    parser.add_argument(
        "--year",
        type=int,
        default=None,
        help="対象年（例: 2026）。省略時はconfig.jsonの値を使用",
    )
    parser.add_argument(
        "--month",
        type=int,
        default=None,
        help="対象月（例: 1〜12）。省略時はconfig.jsonの値を使用",
    )
    args = parser.parse_args()

    # 設定読み込み
    config = load_config()

    # コマンドライン引数で年月が指定されていたらconfig値を上書き
    if args.year is not None:
        config["target_year"] = args.year
    if args.month is not None:
        if not (1 <= args.month <= 12):
            print(f"❌ 月は1〜12で指定してください（指定値: {args.month}）")
            sys.exit(1)
        config["target_month"] = args.month

    debug_port = config.get("chrome_debug_port", 9222)

    logger.info("=" * 50)
    logger.info("報告書PDF自動出力ツール 起動")
    logger.info("=" * 50)

    # リモートデバッグ付きChromeの接続確認
    if not check_chrome_debuggable(debug_port):
        print()
        print("❌ リモートデバッグ付きChromeが見つかりません。")
        print()
        print("以下の手順で準備してください:")
        print("  1. Chromeが開いていたら完全に閉じる (Cmd+Q)")
        print("  2. start.command をダブルクリック、または以下を実行:")
        print(f'     /Applications/Google\\ Chrome.app/Contents/MacOS/Google\\ Chrome'
              f' --remote-debugging-port={debug_port} &')
        print("  3. Chromeでログインし「介護報告書一括印刷」画面を開く")
        print("  4. このツールを再実行")
        sys.exit(1)

    # 既に開いているChromeに接続
    logger.info("Chromeに接続中...")
    driver = None
    try:
        driver = create_driver(config)
        logger.info("Chrome接続完了")

        # 現在のページを確認
        current_url = driver.current_url
        logger.info(f"現在のページ: {current_url}")

        print()
        print("Chromeに接続しました。")
        print(f"現在のページ: {current_url}")
        print()
        print("「介護報告書一括印刷」画面が表示されていることを確認してください。")
        input("準備ができたらEnterキーを押してください...")
        print()

        if args.dry_run:
            run_dry_run(driver, config)
        else:
            run_main(driver, config)

    except WebDriverException as e:
        logger.error(f"Chrome接続エラー: {e}")
        print(f"\n❌ Chrome接続エラー: {e}")
        print("Chromeがリモートデバッグモードで起動しているか確認してください。")
        sys.exit(1)

    except KeyboardInterrupt:
        logger.info("ユーザーによる中断")
        print("\n処理を中断しました。progress.json に進捗が保存されています。")

    except Exception as e:
        logger.error(f"予期せぬエラー: {e}")
        logger.debug(traceback.format_exc())
        print(f"\n❌ 予期せぬエラー: {e}")

    finally:
        # 接続を切断するだけ（Chromeは閉じない）
        if driver:
            try:
                driver.quit()
            except Exception:
                pass
        logger.info("ツール終了")


if __name__ == "__main__":
    main()
