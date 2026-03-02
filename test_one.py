#!/usr/bin/env python3
"""
テストスクリプト: ケアマネ1名分のPDFダウンロードを実行して動作確認する。
input() 待ちを省略し、自動で処理する。
"""
import os
import sys
import time

# main.py と同じディレクトリにいることを確認
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, BASE_DIR)

from main import (
    load_config,
    setup_logging,
    create_driver,
    navigate_to_page,
    set_search_conditions,
    get_care_manager_list,
    get_date_range,
    make_output_dir,
    process_care_manager,
    logger,
)


def main():
    config = load_config()

    print("=" * 50)
    print("  テスト: ケアマネ1名分のPDFダウンロード")
    print("=" * 50)

    # Chrome接続
    print("\nChromeに接続中...")
    driver = create_driver(config)
    print(f"接続完了: {driver.current_url}")

    try:
        year = config["target_year"]
        month = config["target_month"]
        from_date, to_date = get_date_range(year, month)
        wait_timeout = config["wait_time"]["element_timeout"]

        # 一括印刷画面へ移動
        navigate_to_page(driver, config)
        print(f"一括印刷画面へ移動: {driver.current_url}")

        # 検索条件設定
        set_search_conditions(driver, config, from_date, to_date, wait_timeout)
        print("検索条件設定完了")

        # ケアマネリスト取得
        care_managers = get_care_manager_list(driver, config, wait_timeout)
        print(f"\nケアマネリスト: 全{len(care_managers)}名")
        for i, cm in enumerate(care_managers[:5], 1):
            print(f"  {i}. {cm['name']} (index={cm['index']})")
        if len(care_managers) > 5:
            print(f"  ... 他{len(care_managers) - 5}名")

        if not care_managers:
            print("\nケアマネリストが空です。テスト中止。")
            return

        # 最初の1名でテスト
        test_cm = care_managers[0]
        print(f"\n--- テスト対象: {test_cm['name']}様 ---")
        print("PDFダウンロードを開始します...")

        output_dir = make_output_dir(config, year, month)
        start_time = time.time()

        result_status, result_reason = process_care_manager(
            driver, config, test_cm, 1, 1, from_date, to_date, output_dir
        )

        elapsed = time.time() - start_time

        print(f"\n--- テスト結果 ---")
        print(f"ステータス: {result_status}")
        if result_reason:
            print(f"理由: {result_reason}")
        print(f"所要時間: {int(elapsed)}秒")
        print(f"出力先: {output_dir}")

        # 出力先のファイルを確認
        if os.path.exists(output_dir):
            files = os.listdir(output_dir)
            pdf_files = [f for f in files if f.endswith('.pdf')]
            if pdf_files:
                print(f"ダウンロードされたPDF:")
                for f in pdf_files:
                    filepath = os.path.join(output_dir, f)
                    size_kb = os.path.getsize(filepath) / 1024
                    print(f"  {f} ({size_kb:.1f} KB)")
            else:
                print("PDFファイルが見つかりません")

    except Exception as e:
        print(f"\nエラー: {e}")
        import traceback
        traceback.print_exc()

    finally:
        try:
            driver.quit()
        except Exception:
            pass


if __name__ == "__main__":
    main()
