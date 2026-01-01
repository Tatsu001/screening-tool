#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
投資管理テンプレート - スクリーニングシート自動更新ツール
Version: 3.4.0
"""

import yfinance as yf
import openpyxl
from openpyxl.styles import PatternFill, Border, Side, Alignment, Font
import sys
import time
import os
import glob

# Tkinterのインポート（GUIファイル選択用）
try:
    import tkinter as tk
    from tkinter import filedialog
    GUI_AVAILABLE = True
except ImportError:
    GUI_AVAILABLE = False

# 色の定義
PORTFOLIO_ALERT_COLOR = 'FFA500'  # オレンジ色（ポートフォリオアラート用）

def get_stock_data(ticker_code):
    """
    yfinanceで株価データを取得
    
    Args:
        ticker_code: 銘柄コード（4桁）
    
    Returns:
        dict: 株価データ
    """
    try:
        # 日本株は .T を付ける
        ticker = f"{ticker_code}.T"
        stock = yf.Ticker(ticker)
        info = stock.info
        
        # 基本情報
        name = info.get('longName', info.get('shortName', '-'))
        market_cap = info.get('marketCap')
        trailing_pe = info.get('trailingPE')
        price_to_book = info.get('priceToBook')
        
        # 自己資本比率を計算
        equity_ratio = None
        total_equity = info.get('totalStockholderEquity')
        total_assets = info.get('totalAssets')
        
        if total_equity and total_assets and total_assets != 0:
            equity_ratio = (total_equity / total_assets) * 100
        
        # ROE
        return_on_equity = info.get('returnOnEquity')
        if return_on_equity is not None:
            return_on_equity = return_on_equity * 100  # パーセント変換
        
        # 売上成長率
        revenue_growth = info.get('revenueGrowth')
        if revenue_growth is not None:
            revenue_growth = revenue_growth * 100  # パーセント変換
        
        # 過去データから売買代金を計算
        hist = stock.history(period='5d')
        trading_value = None
        
        if not hist.empty and 'Volume' in hist.columns and 'Close' in hist.columns:
            # 最新5日間の平均売買代金
            hist['Value'] = hist['Volume'] * hist['Close']
            trading_value = hist['Value'].mean()
        
        return {
            'name': name,
            'market_cap': market_cap,
            'equity_ratio': equity_ratio,
            'trading_value': trading_value,
            'trailing_pe': trailing_pe,
            'price_to_book': price_to_book,
            'return_on_equity': return_on_equity,
            'revenue_growth': revenue_growth,
        }
    
    except Exception as e:
        print(f"  エラー: {str(e)}")
        return None

def format_value(value, format_type='number', decimals=1):
    """
    値をフォーマット
    
    Args:
        value: フォーマットする値
        format_type: フォーマットタイプ（number, percent, currency）
        decimals: 小数点以下の桁数
    
    Returns:
        フォーマットされた値、またはNoneの場合は'-'
    """
    if value is None:
        return '-'
    
    try:
        if format_type == 'number':
            return round(value, decimals)
        elif format_type == 'percent':
            return round(value, decimals)
        elif format_type == 'currency':
            return round(value, 0)
        else:
            return value
    except:
        return '-'

def get_stocks_from_sheet(wb, sheet_name):
    """
    指定したシートから銘柄コードリストを取得
    
    Args:
        wb: openpyxlのワークブック
        sheet_name: シート名
    
    Returns:
        list: 銘柄コードのリスト
    """
    if sheet_name not in wb.sheetnames:
        return []
    
    ws = wb[sheet_name]
    stock_codes = []
    
    # A列の2行目以降から銘柄コードを取得
    for row in range(2, 100):  # 最大98銘柄
        code = ws[f'A{row}'].value
        if code and str(code).strip():
            stock_codes.append(str(code).strip())
        elif not code:
            # 空欄が出たら終了
            break
    
    return stock_codes

def get_portfolio_stocks(wb):
    """
    ポートフォリオシートから保有銘柄のコードリストを取得
    
    Args:
        wb: openpyxlのワークブック
    
    Returns:
        set: 銘柄コードのセット
    """
    if 'ポートフォリオ' not in wb.sheetnames:
        return set()
    
    ws = wb['ポートフォリオ']
    stock_codes = set()
    
    # A列の7行目以降から銘柄コードを取得
    for row in range(7, 100):
        code = ws[f'A{row}'].value
        if code and str(code).strip():
            stock_codes.add(str(code).strip())
    
    return stock_codes

def update_screening_sheet(filepath, stock_codes, market_map):
    """
    スクリーニングシートを更新
    
    Args:
        filepath: Excelファイルパス
        stock_codes: 更新する銘柄コードのリスト
        market_map: 銘柄コードと市場区分のマッピング辞書
    """
    print(f"\n📊 ファイルを読み込み中: {filepath}")
    
    try:
        wb = openpyxl.load_workbook(filepath)
    except FileNotFoundError:
        print(f"❌ エラー: ファイルが見つかりません - {filepath}")
        sys.exit(1)
    except Exception as e:
        print(f"❌ エラー: ファイルの読み込みに失敗 - {str(e)}")
        sys.exit(1)
    
    if 'スクリーニング' not in wb.sheetnames:
        print("❌ エラー: 'スクリーニング'シートが見つかりません")
        sys.exit(1)
    
    ws = wb['スクリーニング']
    
    # ポートフォリオの保有銘柄を取得
    portfolio_stocks = get_portfolio_stocks(wb)
    print(f"\n🔍 ポートフォリオ保有銘柄: {len(portfolio_stocks)}銘柄")
    if portfolio_stocks:
        print(f"   {', '.join(sorted(portfolio_stocks))}")
    
    # 前回のスクリーニングシートから既存銘柄とI列以降のデータを保存
    print(f"\n📋 既存データを読み込み中...")
    existing_data = {}  # {銘柄コード: {row_data: I列以降のデータ}}
    
    for row in range(6, 21):  # 6～20行目
        code = ws[f'A{row}'].value
        if code and str(code).strip():
            code = str(code).strip()
            # I列以降（9列目以降）のデータを保存
            row_data = {}
            for col in range(9, 25):  # I列(9)～X列(24)
                cell = ws.cell(row=row, column=col)
                row_data[col] = {
                    'value': cell.value,
                    'fill': cell.fill.copy() if cell.fill else None,
                    'font': cell.font.copy() if cell.font else None,
                    'alignment': cell.alignment.copy() if cell.alignment else None,
                    'border': cell.border.copy() if cell.border else None,
                    'number_format': cell.number_format,
                }
            existing_data[code] = row_data
            print(f"   {code}: I列以降のデータを保存")
    
    # データの最終行を見つける（新規行のテンプレート用）
    template_row = None
    for row in range(6, 21):
        code = ws[f'A{row}'].value
        if not code or not str(code).strip():
            template_row = row
            break
    if template_row is None:
        template_row = 21  # 見つからない場合は21行目
    
    print(f"\n📝 テンプレート行: {template_row}行目")
    
    # 統合リストを作成
    stock_codes_set = set(stock_codes)
    portfolio_only = portfolio_stocks - stock_codes_set
    
    unified_list = list(stock_codes) + list(portfolio_only)
    
    print(f"\n📊 統合リスト: {len(unified_list)}銘柄")
    print(f"   スクリーニング銘柄: {len(stock_codes)}銘柄")
    print(f"   ポートフォリオのみ: {len(portfolio_only)}銘柄")
    
    # スタイル定義
    alert_fill = PatternFill(start_color=PORTFOLIO_ALERT_COLOR, end_color=PORTFOLIO_ALERT_COLOR, fill_type='solid')
    thin_border = Border(
        left=Side(style='thin'),
        right=Side(style='thin'),
        top=Side(style='thin'),
        bottom=Side(style='thin')
    )
    center_align = Alignment(horizontal='center', vertical='center', wrap_text=True)
    
    # A～H列とJ～K列のみクリア（I列は触らない）
    print(f"\n🧹 A～H列、J～K列をクリア中...")
    for row in range(6, 21):
        for col in range(1, 9):  # A列(1)～H列(8)
            cell = ws.cell(row=row, column=col)
            cell.value = None
            cell.fill = openpyxl.styles.PatternFill(fill_type=None)
        for col in range(10, 12):  # J列(10)～K列(11)
            cell = ws.cell(row=row, column=col)
            cell.value = None
            cell.fill = openpyxl.styles.PatternFill(fill_type=None)
    
    print(f"\n📡 株価データを取得中...")
    print("=" * 60)
    
    # データ開始行
    current_row = 6
    portfolio_alerts = []
    
    # 統合リストの各銘柄を処理
    for idx, code in enumerate(unified_list, start=1):
        code = str(code).strip()
        
        print(f"\n[{idx}/{len(unified_list)}] {code}")
        
        # ポートフォリオにあるが今回のリストにない = オレンジ色
        is_portfolio_alert = code in portfolio_only
        if is_portfolio_alert:
            print(f"  ⚠️  ポートフォリオ保有中（スクリーニング対象外）")
            portfolio_alerts.append(code)
        
        # yfinanceでデータ取得
        print(f"  取得中...", end=" ")
        data = get_stock_data(code)
        
        if data is None:
            print("スキップ")
            current_row += 1
            continue
        
        print("✓")
        
        # 市場区分を取得（market_mapから）
        market = market_map.get(code, '')
        if market:
            print(f"  市場区分: {market}")
        
        # 新規銘柄の場合、テンプレート行から書式・入力規則をコピー
        is_new_stock = code not in existing_data
        if is_new_stock:
            print(f"  📋 新規銘柄: テンプレート行から書式をコピー")
            # I列以降の書式・入力規則をコピー（値はコピーしない）
            for col in range(9, 25):  # I列(9)～X列(24)
                template_cell = ws.cell(row=template_row, column=col)
                target_cell = ws.cell(row=current_row, column=col)
                
                # 値はコピーしない（空欄のまま）
                target_cell.value = None
                
                # 書式をコピー
                if template_cell.fill:
                    target_cell.fill = template_cell.fill.copy()
                if template_cell.font:
                    target_cell.font = template_cell.font.copy()
                if template_cell.alignment:
                    target_cell.alignment = template_cell.alignment.copy()
                if template_cell.border:
                    target_cell.border = template_cell.border.copy()
                if template_cell.number_format:
                    target_cell.number_format = template_cell.number_format
        
        # A～H列を書き込み（オレンジ色はポートフォリオアラートのみ）
        row = current_row
        
        # A列: 銘柄コード
        ws[f'A{row}'] = code
        if is_portfolio_alert:
            ws[f'A{row}'].fill = alert_fill
        ws[f'A{row}'].alignment = center_align
        ws[f'A{row}'].border = thin_border
        
        # B列: 銘柄名
        name = data['name'] if data['name'] and data['name'] != '-' else '-'
        ws[f'B{row}'] = name
        if is_portfolio_alert:
            ws[f'B{row}'].fill = alert_fill
        ws[f'B{row}'].alignment = center_align
        ws[f'B{row}'].border = thin_border
        
        # C列: 市場区分（market_mapから取得、空欄の場合もあり）
        ws[f'C{row}'] = market
        if is_portfolio_alert:
            ws[f'C{row}'].fill = alert_fill
        ws[f'C{row}'].alignment = center_align
        ws[f'C{row}'].border = thin_border
        
        # D列: 時価総額
        market_cap = format_value(data['market_cap'] / 100000000 if data['market_cap'] else None, 'currency')
        ws[f'D{row}'] = market_cap
        if market_cap != '-':
            ws[f'D{row}'].number_format = '#,##0'
        if is_portfolio_alert:
            ws[f'D{row}'].fill = alert_fill
        ws[f'D{row}'].alignment = center_align
        ws[f'D{row}'].border = thin_border
        
        # E列: 自己資本比率
        equity_ratio = format_value(data['equity_ratio'], 'number', 1)
        ws[f'E{row}'] = equity_ratio
        if equity_ratio != '-':
            ws[f'E{row}'].number_format = '0.0'
        if is_portfolio_alert:
            ws[f'E{row}'].fill = alert_fill
        ws[f'E{row}'].alignment = center_align
        ws[f'E{row}'].border = thin_border
        
        # F列: 売買代金
        trading_value = format_value(data['trading_value'], 'currency')
        ws[f'F{row}'] = trading_value
        if trading_value != '-':
            ws[f'F{row}'].number_format = '#,##0'
        if is_portfolio_alert:
            ws[f'F{row}'].fill = alert_fill
        ws[f'F{row}'].alignment = center_align
        ws[f'F{row}'].border = thin_border
        
        # G列: PER
        per = format_value(data['trailing_pe'], 'number', 1)
        ws[f'G{row}'] = per
        if per != '-':
            ws[f'G{row}'].number_format = '0.0'
        if is_portfolio_alert:
            ws[f'G{row}'].fill = alert_fill
        ws[f'G{row}'].alignment = center_align
        ws[f'G{row}'].border = thin_border
        
        # H列: PBR
        pbr = format_value(data['price_to_book'], 'number', 1)
        ws[f'H{row}'] = pbr
        if pbr != '-':
            ws[f'H{row}'].number_format = '0.0'
        if is_portfolio_alert:
            ws[f'H{row}'].fill = alert_fill
        ws[f'H{row}'].alignment = center_align
        ws[f'H{row}'].border = thin_border
        
        # I列: バリュースコア（数式 - 触らない）
        
        # J列: 売上成長率（自動取得）
        revenue_growth = format_value(data['revenue_growth'], 'percent', 1)
        ws[f'J{row}'] = revenue_growth
        if revenue_growth != '-':
            ws[f'J{row}'].number_format = '0.0'
        if is_portfolio_alert:
            ws[f'J{row}'].fill = alert_fill
        ws[f'J{row}'].alignment = center_align
        ws[f'J{row}'].border = thin_border
        
        # K列: ROE（自動取得）
        roe = format_value(data['return_on_equity'], 'percent', 1)
        ws[f'K{row}'] = roe
        if roe != '-':
            ws[f'K{row}'].number_format = '0.0'
        if is_portfolio_alert:
            ws[f'K{row}'].fill = alert_fill
        ws[f'K{row}'].alignment = center_align
        ws[f'K{row}'].border = thin_border
        
        # I列以降: 既存データがあれば復元（数式・手動入力を保持）
        if code in existing_data:
            print(f"  📋 I列以降のデータを復元")
            for col, cell_data in existing_data[code].items():
                cell = ws.cell(row=row, column=col)
                cell.value = cell_data['value']
                if cell_data['fill']:
                    cell.fill = cell_data['fill']
                if cell_data['font']:
                    cell.font = cell_data['font']
                if cell_data['alignment']:
                    cell.alignment = cell_data['alignment']
                if cell_data['border']:
                    cell.border = cell_data['border']
                if cell_data['number_format']:
                    cell.number_format = cell_data['number_format']
        
        current_row += 1
        
        # API制限を避けるため少し待機
        time.sleep(0.5)
    
    # ファイルを保存
    print("\n" + "=" * 60)
    print(f"💾 ファイルを保存中...")
    
    try:
        wb.save(filepath)
        print(f"✅ 保存完了: {filepath}")
    except Exception as e:
        print(f"❌ エラー: ファイルの保存に失敗 - {str(e)}")
        sys.exit(1)
    
    # サマリー表示
    print("\n" + "=" * 60)
    print("📊 更新サマリー")
    print("=" * 60)
    print(f"更新銘柄数: {len(unified_list)}銘柄")
    print(f"  - スクリーニング銘柄: {len(stock_codes)}銘柄")
    print(f"  - ポートフォリオのみ: {len(portfolio_only)}銘柄")
    
    if portfolio_alerts:
        print(f"\n⚠️  ポートフォリオ保有中（スクリーニング対象外）:")
        for code in portfolio_alerts:
            print(f"   - {code}")
        print(f"\n注意: これらの銘柄は売却を検討してください。")
    
    print("\n✅ スクリーニングシート更新完了!")

def main():
    """
    メイン関数
    """
    print("=" * 60)
    print("📊 投資管理テンプレート - スクリーニングシート自動更新")
    print("=" * 60)
    
    filepath = None
    
    # GUIでファイル選択を試みる
    if GUI_AVAILABLE:
        try:
            print("\n📁 ファイル選択ダイアログを開きます...")
            root = tk.Tk()
            root.withdraw()
            
            filepath = filedialog.askopenfilename(
                title="Excelファイルを選択",
                filetypes=[
                    ("Excelファイル", "*.xlsx"),
                    ("すべてのファイル", "*.*")
                ]
            )
            
            root.destroy()
            
            if filepath:
                print(f"✅ 選択されたファイル: {filepath}")
            else:
                print("❌ ファイルが選択されませんでした")
        except Exception as e:
            print(f"⚠️  GUI選択に失敗: {str(e)}")
            if filepath:
                print("❌ ファイルが選択されませんでした")
        
        # GUIが使えないか、キャンセルされた場合は自動検出
        if not filepath:
            print("\n📁 Excelファイルを自動検出します...")
            
            # 候補となるファイル名
            candidates = [
                'investment_template.xlsx',
                '投資管理テンプレート.xlsx',
                '投資管理テンプレート_配列数式版.xlsx',
            ]
            
            # カレントディレクトリで検索
            for candidate in candidates:
                if os.path.exists(candidate):
                    filepath = candidate
                    print(f"✅ 発見: {filepath}")
                    break
            
            # 見つからない場合、xlsxファイルを全て表示
            if not filepath:
                xlsx_files = glob.glob('*.xlsx')
                if xlsx_files:
                    print("\n以下のExcelファイルが見つかりました:")
                    for i, f in enumerate(xlsx_files, 1):
                        print(f"  {i}. {f}")
                    
                    print("\n使用するファイル番号を入力してください:")
                    try:
                        choice = int(input("番号: ").strip())
                        if 1 <= choice <= len(xlsx_files):
                            filepath = xlsx_files[choice - 1]
                            print(f"✅ 選択: {filepath}")
                        else:
                            print("❌ エラー: 無効な番号です")
                            input("\nEnterキーで終了...")
                            sys.exit(1)
                    except (ValueError, EOFError):
                        print("❌ エラー: 無効な入力です")
                        input("\nEnterキーで終了...")
                        sys.exit(1)
                else:
                    print("\n❌ エラー: Excelファイルが見つかりません")
                    print("\n以下のいずれかのファイルを同じフォルダに配置してください:")
                    print("  - investment_template.xlsx")
                    print("  - 投資管理テンプレート.xlsx")
                    input("\nEnterキーで終了...")
                    sys.exit(1)
    
    # ファイルの存在確認
    if not os.path.exists(filepath):
        print(f"\n❌ エラー: ファイルが見つかりません - {filepath}")
        input("\nEnterキーで終了...")
        sys.exit(1)
    
    # Excelファイルを開いて銘柄リストを取得
    print(f"\n📊 ファイルを読み込み中: {filepath}")
    
    try:
        wb = openpyxl.load_workbook(filepath)
    except Exception as e:
        print(f"❌ エラー: ファイルの読み込みに失敗 - {str(e)}")
        input("\nEnterキーで終了...")
        sys.exit(1)
    
    # 各シートから銘柄コードを取得
    growth_stocks = get_stocks_from_sheet(wb, '銘柄スクリーニング（グロース）')
    prime_stocks = get_stocks_from_sheet(wb, '銘柄スクリーニング（プライム）')
    other_stocks = get_stocks_from_sheet(wb, 'スクリーニング銘柄')
    
    wb.close()
    
    # 市場区分のマッピングを作成
    market_map = {}
    
    # グロースシートの銘柄
    for code in growth_stocks:
        market_map[code] = 'グロース'
    
    # プライムシートの銘柄（重複チェック）
    for code in prime_stocks:
        if code in market_map:
            # 重複の場合は空欄
            market_map[code] = ''
        else:
            market_map[code] = 'プライム'
    
    # スクリーニング銘柄シートの銘柄（重複チェック）
    for code in other_stocks:
        if code in market_map:
            # 重複の場合は空欄
            market_map[code] = ''
        else:
            market_map[code] = ''  # 元々空欄
    
    # 統合リスト作成
    all_stocks = set(growth_stocks + prime_stocks + other_stocks)
    stock_codes = list(all_stocks)
    
    if not stock_codes:
        print("\n❌ エラー: 銘柄コードが入力されていません")
        print("\n手順:")
        print("1. Excelファイルを開く")
        print("2. 以下のいずれかのシートのA列（2行目以降）に銘柄コードを入力")
        print("   - 銘柄スクリーニング（グロース）")
        print("   - 銘柄スクリーニング（プライム）")
        print("   - スクリーニング銘柄")
        print("3. 保存してから再実行")
        input("\nEnterキーで終了...")
        sys.exit(1)
    
    # 情報表示
    print(f"\n📊 読み込んだ銘柄:")
    print(f"   銘柄スクリーニング（グロース）: {len(growth_stocks)}銘柄")
    if growth_stocks:
        print(f"     {', '.join(growth_stocks)}")
    
    print(f"   銘柄スクリーニング（プライム）: {len(prime_stocks)}銘柄")
    if prime_stocks:
        print(f"     {', '.join(prime_stocks)}")
    
    print(f"   スクリーニング銘柄: {len(other_stocks)}銘柄")
    if other_stocks:
        print(f"     {', '.join(other_stocks)}")
    
    # 重複チェック
    duplicates = []
    checked = set()
    for code in growth_stocks + prime_stocks + other_stocks:
        if code in checked:
            if code not in duplicates:
                duplicates.append(code)
        else:
            checked.add(code)
    
    if duplicates:
        print(f"\n⚠️  重複銘柄（市場区分: 空欄）:")
        print(f"     {', '.join(duplicates)}")
    
    # 確認
    print(f"\n✅ 合計 {len(stock_codes)}銘柄を更新します")
    print()
    
    try:
        confirm = input("続行しますか？ (y/N): ").strip().lower()
    except EOFError:
        confirm = 'n'
    
    if confirm != 'y':
        print("\n❌ キャンセルされました")
        input("\nEnterキーで終了...")
        sys.exit(0)
    
    # スクリーニングシートを更新
    update_screening_sheet(filepath, stock_codes, market_map)
    
    # 終了
    input("\nEnterキーで終了...")

if __name__ == '__main__':
    try:
        main()
    except KeyboardInterrupt:
        print("\n\n❌ 中断されました")
        sys.exit(1)
    except Exception as e:
        print(f"\n❌ 予期しないエラー: {str(e)}")
        import traceback
        traceback.print_exc()
        input("\nEnterキーで終了...")
        sys.exit(1)