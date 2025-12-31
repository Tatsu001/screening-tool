#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
投資管理テンプレート - スクリーニングシート自動更新スクリプト

使い方:
    python update_screening.py 投資管理テンプレート.xlsx

機能:
    1. yfinanceで株価・財務データを取得
    2. スクリーニングシートのみ上書き
    3. ポートフォリオに残っている銘柄は保持（背景色でアラート）
    4. その他のシートは変更なし
"""

import sys
import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from datetime import datetime
import time
import os
import glob

# tkinterをインポート（ファイル選択ダイアログ用）
try:
    import tkinter as tk
    from tkinter import filedialog
    HAS_TKINTER = True
except ImportError:
    HAS_TKINTER = False
    print("⚠️  tkinterが利用できません。ファイル選択ダイアログは使用できません。")

# yfinanceのインストール確認
try:
    import yfinance as yf
except ImportError:
    print("yfinanceがインストールされていません。")
    print("以下のコマンドでインストールしてください:")
    print("  pip install yfinance")
    input("\nEnterキーで終了...")
    sys.exit(1)

# 色定義
HEADER_COLOR = "2C3E50"
SUBHEADER_COLOR = "34495E"
INPUT_COLOR = "FFF9E6"
WHITE = "FFFFFF"
SUCCESS_COLOR = "D5F4E6"
WARNING_COLOR = "FCF3CF"
DANGER_COLOR = "FADBD8"
PORTFOLIO_ALERT_COLOR = "FFE5CC"  # ポートフォリオ銘柄アラート色（オレンジ）

def select_excel_file():
    """
    GUIファイル選択ダイアログでExcelファイルを選択
    
    Returns:
        str: 選択されたファイルパス（キャンセル時はNone）
    """
    if not HAS_TKINTER:
        return None
    
    # tkinterのルートウィンドウを作成（非表示）
    root = tk.Tk()
    root.withdraw()
    root.attributes('-topmost', True)
    
    # ファイル選択ダイアログを表示
    filepath = filedialog.askopenfilename(
        title='Excelファイルを選択してください',
        filetypes=[
            ('Excel files', '*.xlsx'),
            ('All files', '*.*')
        ],
        initialdir=os.getcwd()
    )
    
    root.destroy()
    
    return filepath if filepath else None

# グロース市場の銘柄リスト（代表的な銘柄）
GROWTH_MARKET_CODES = {
    '4478', '4755', '4477', '4481', '4486', '4488', '3681', '3696',
    '7047', '7048', '7049', '6070', '6098', '6177', '6178', '6180',
    '4385', '4386', '4431', '4433', '4434', '4435', '4436', '4479',
    '2158', '2326', '2379', '2427', '2428', '3923', '3924', '3928',
    '4368', '4371', '4374', '4375', '4376', '4378', '4382', '4384',
}

# スタンダード市場の銘柄リスト（代表的な銘柄）
STANDARD_MARKET_CODES = {
    '1515', '1518', '1719', '1720', '1721', '1766', '1770', '1780',
    '5401', '5410', '5411', '5444', '5445', '5449', '5451', '5471',
}

def get_market_category(info, ticker_code):
    """
    市場区分を取得
    
    Args:
        info: yfinanceのinfo辞書
        ticker_code: 銘柄コード（4桁）
    
    Returns:
        str: プライム/スタンダード/グロース
    """
    # まず銘柄コードから判定（最も確実）
    if ticker_code in GROWTH_MARKET_CODES:
        return 'グロース'
    
    if ticker_code in STANDARD_MARKET_CODES:
        return 'スタンダード'
    
    # yfinanceのデータから判定を試みる
    exchange = str(info.get('exchange', '')).upper()
    quote_type = str(info.get('quoteType', '')).upper()
    long_name = str(info.get('longName', '')).lower()
    
    # グロース市場のキーワード判定
    if any(keyword in long_name for keyword in ['growth', 'mothers', 'jasdaq growth']):
        return 'グロース'
    
    # 市場情報から判定
    if 'market' in info:
        market_info = str(info.get('market', '')).lower()
        if 'growth' in market_info or 'mothers' in market_info:
            return 'グロース'
        elif 'standard' in market_info:
            return 'スタンダード'
        elif 'prime' in market_info:
            return 'プライム'
    
    # デフォルトはプライム（主要企業は大抵プライム）
    return 'プライム'

def get_stock_data(ticker_code):
    """
    yfinanceで株価・財務データを取得
    
    Args:
        ticker_code: 銘柄コード（例: 7203）
    
    Returns:
        dict: 取得したデータ（データがない場合は'-'）
    """
    try:
        # 日本株は .T を付ける
        ticker = f"{ticker_code}.T"
        stock = yf.Ticker(ticker)
        info = stock.info
        
        # 市場区分を取得（銘柄コードを渡す）
        market = get_market_category(info, ticker_code)
        
        # 自己資本比率を計算
        equity_ratio = None
        if info.get('totalStockholderEquity') and info.get('totalAssets'):
            total_equity = info.get('totalStockholderEquity')
            total_assets = info.get('totalAssets')
            if total_assets > 0:
                equity_ratio = (total_equity / total_assets) * 100
        
        # 売買代金を計算（億円単位）
        trading_value = None
        if info.get('averageVolume') and info.get('currentPrice'):
            avg_volume = info.get('averageVolume')
            current_price = info.get('currentPrice')
            # 出来高 × 株価 ÷ 100,000,000 = 億円
            trading_value = (avg_volume * current_price) / 100000000
        
        # ROE
        roe = info.get('returnOnEquity')
        if roe is not None:
            roe = roe * 100  # パーセント表記
        
        # 売上成長率
        revenue_growth = info.get('revenueGrowth')
        if revenue_growth is not None:
            revenue_growth = revenue_growth * 100  # パーセント表記
        
        # データを辞書形式で返す（Noneの場合は'-'に変換）
        data = {
            'name': info.get('longName', info.get('shortName', '-')),
            'market': market,
            'market_cap': info.get('marketCap'),
            'equity_ratio': equity_ratio,
            'trading_value': trading_value,
            'trailing_pe': info.get('trailingPE'),
            'price_to_book': info.get('priceToBook'),
            'return_on_equity': roe,
            'revenue_growth': revenue_growth,
        }
        
        return data
        
    except Exception as e:
        print(f"  ⚠️  {ticker_code}: データ取得エラー - {str(e)}")
        return None



def format_value(value, format_type='number', decimals=1):
    """
    値をフォーマット（Noneの場合は'-'を返す）
    
    Args:
        value: フォーマットする値
        format_type: 'number', 'percent', 'currency'
        decimals: 小数点以下の桁数
    
    Returns:
        フォーマットされた値または'-'
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

def get_portfolio_stocks(wb):
    """
    ポートフォリオシートから保有銘柄のコードリストを取得
    
    Args:
        wb: openpyxlのワークブック
    
    Returns:
        set: 保有銘柄コードのセット
    """
    portfolio_stocks = set()
    
    if 'ポートフォリオ' not in wb.sheetnames:
        return portfolio_stocks
    
    ws = wb['ポートフォリオ']
    
    # 7行目から11行目まで（データ行）
    for row in range(7, 12):
        code = ws[f'A{row}'].value
        if code and str(code).strip():
            portfolio_stocks.add(str(code).strip())
    
    return portfolio_stocks

def update_screening_sheet(filepath, stock_codes):
    """
    スクリーニングシートを更新
    
    Args:
        filepath: Excelファイルパス
        stock_codes: 更新する銘柄コードのリスト
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
    
    # スタイル定義
    input_fill = PatternFill(start_color=INPUT_COLOR, end_color=INPUT_COLOR, fill_type='solid')
    alert_fill = PatternFill(start_color=PORTFOLIO_ALERT_COLOR, end_color=PORTFOLIO_ALERT_COLOR, fill_type='solid')
    thin_border = Border(
        left=Side(style='thin'),
        right=Side(style='thin'),
        top=Side(style='thin'),
        bottom=Side(style='thin')
    )
    center_align = Alignment(horizontal='center', vertical='center', wrap_text=True)
    
    print(f"\n📡 株価データを取得中...")
    print("=" * 60)
    
    # データ開始行（6行目から）
    start_row = 6
    current_row = start_row
    
    # 既存データをクリア（6行目以降）
    for row in range(6, 21):
        for col in range(1, 25):
            cell = ws.cell(row=row, column=col)
            cell.value = None
    
    # 各銘柄のデータを取得して書き込み
    portfolio_alerts = []
    
    for idx, code in enumerate(stock_codes, start=1):
        code = str(code).strip()
        
        print(f"\n[{idx}/{len(stock_codes)}] {code}")
        
        # ポートフォリオ保有銘柄かチェック
        is_portfolio_stock = code in portfolio_stocks
        if is_portfolio_stock:
            print(f"  ⚠️  ポートフォリオ保有中")
            portfolio_alerts.append(code)
        
        # yfinanceでデータ取得
        print(f"  取得中...", end=" ")
        data = get_stock_data(code)
        
        if data is None:
            print("スキップ")
            # データ取得失敗でも行は進める
            current_row += 1
            continue
        
        print("✓")
        
        # データを書き込み
        row = current_row
        
        # A列: 銘柄コード
        ws[f'A{row}'] = code
        ws[f'A{row}'].fill = alert_fill if is_portfolio_stock else input_fill
        ws[f'A{row}'].alignment = center_align
        ws[f'A{row}'].border = thin_border
        
        # B列: 銘柄名
        name = data['name'] if data['name'] and data['name'] != '-' else '-'
        ws[f'B{row}'] = name
        ws[f'B{row}'].fill = alert_fill if is_portfolio_stock else input_fill
        ws[f'B{row}'].alignment = center_align
        ws[f'B{row}'].border = thin_border
        
        # C列: 市場区分
        ws[f'C{row}'] = data.get('market', 'プライム')
        ws[f'C{row}'].fill = alert_fill if is_portfolio_stock else input_fill
        ws[f'C{row}'].alignment = center_align
        ws[f'C{row}'].border = thin_border
        
        # D列: 時価総額
        market_cap = format_value(data['market_cap'] / 100000000 if data['market_cap'] else None, 'currency')
        ws[f'D{row}'] = market_cap
        if market_cap != '-':
            ws[f'D{row}'].number_format = '#,##0'
        ws[f'D{row}'].fill = alert_fill if is_portfolio_stock else input_fill
        ws[f'D{row}'].alignment = center_align
        ws[f'D{row}'].border = thin_border
        
        # E列: 自己資本比率
        equity_ratio = format_value(data['equity_ratio'], 'number', 1)
        ws[f'E{row}'] = equity_ratio
        if equity_ratio != '-':
            ws[f'E{row}'].number_format = '0.0'
        ws[f'E{row}'].fill = alert_fill if is_portfolio_stock else input_fill
        ws[f'E{row}'].alignment = center_align
        ws[f'E{row}'].border = thin_border
        
        # F列: 売買代金
        trading_value = format_value(data['trading_value'], 'currency')
        ws[f'F{row}'] = trading_value
        if trading_value != '-':
            ws[f'F{row}'].number_format = '#,##0'
        ws[f'F{row}'].fill = alert_fill if is_portfolio_stock else input_fill
        ws[f'F{row}'].alignment = center_align
        ws[f'F{row}'].border = thin_border
        
        # G列: PER
        per = format_value(data['trailing_pe'], 'number', 1)
        ws[f'G{row}'] = per
        if per != '-':
            ws[f'G{row}'].number_format = '0.0'
        ws[f'G{row}'].fill = alert_fill if is_portfolio_stock else input_fill
        ws[f'G{row}'].alignment = center_align
        ws[f'G{row}'].border = thin_border
        
        # H列: PBR
        pbr = format_value(data['price_to_book'], 'number', 1)
        ws[f'H{row}'] = pbr
        if pbr != '-':
            ws[f'H{row}'].number_format = '0.0'
        ws[f'H{row}'].fill = alert_fill if is_portfolio_stock else input_fill
        ws[f'H{row}'].alignment = center_align
        ws[f'H{row}'].border = thin_border
        
        # I列: バリュースコア（数式）
        ws[f'I{row}'] = f'=IF(OR(A{row}="",G{row}="",H{row}=""),"",IF(AND(G{row}>=5,G{row}<=10,H{row}>=0.5,H{row}<=0.75),20,IF(AND(G{row}>=5,G{row}<=10,H{row}>0.75,H{row}<=1),18,IF(AND(G{row}>10,G{row}<=15,H{row}>=0.5,H{row}<=0.75),18,IF(AND(G{row}>10,G{row}<=15,H{row}>0.75,H{row}<=1),15,10)))))'
        ws[f'I{row}'].alignment = center_align
        ws[f'I{row}'].border = thin_border
        
        # J列: 売上成長率
        revenue_growth = format_value(data['revenue_growth'], 'percent', 1)
        ws[f'J{row}'] = revenue_growth
        if revenue_growth != '-':
            ws[f'J{row}'].number_format = '0.0'
        ws[f'J{row}'].fill = alert_fill if is_portfolio_stock else input_fill
        ws[f'J{row}'].alignment = center_align
        ws[f'J{row}'].border = thin_border
        
        # K列: ROE
        roe = format_value(data['return_on_equity'], 'percent', 1)
        ws[f'K{row}'] = roe
        if roe != '-':
            ws[f'K{row}'].number_format = '0.0'
        ws[f'K{row}'].fill = alert_fill if is_portfolio_stock else input_fill
        ws[f'K{row}'].alignment = center_align
        ws[f'K{row}'].border = thin_border
        
        # L列: 成長性スコア（数式）
        ws[f'L{row}'] = f'=IF(OR(A{row}="",C{row}="",J{row}=""),"",IF(C{row}="グロース",IF(J{row}>=30,20,IF(J{row}>=20,18,IF(J{row}>=15,15,IF(J{row}>=10,12,10)))),IF(AND(J{row}>=20,K{row}>=15),20,IF(AND(J{row}>=15,K{row}>=12),18,IF(AND(J{row}>=10,K{row}>=10),15,10)))))'
        ws[f'L{row}'].alignment = center_align
        ws[f'L{row}'].border = thin_border
        
        # M-R列: チェックリスト（空欄 - 手動入力）
        for col in range(13, 19):
            ws.cell(row=row, column=col).fill = alert_fill if is_portfolio_stock else input_fill
            ws.cell(row=row, column=col).alignment = center_align
            ws.cell(row=row, column=col).border = thin_border
        
        # S列: 事業性スコア（数式）
        ws[f'S{row}'] = f'=IF(A{row}="","",IF(M{row}="〇",3,IF(M{row}="△",1.5,0))+IF(N{row}="〇",4,IF(N{row}="△",2,0))+IF(O{row}="〇",3,IF(O{row}="△",1.5,0))+IF(P{row}="〇",3,IF(P{row}="△",1.5,0))+IF(Q{row}="〇",4,IF(Q{row}="△",2,0))+IF(R{row}="〇",3,IF(R{row}="△",1.5,0)))'
        ws[f'S{row}'].alignment = center_align
        ws[f'S{row}'].border = thin_border
        
        # T列: トレンドスコア（空欄 - 手動入力）
        ws[f'T{row}'].fill = alert_fill if is_portfolio_stock else input_fill
        ws[f'T{row}'].alignment = center_align
        ws[f'T{row}'].border = thin_border
        
        # U列: 総合スコア（数式）
        ws[f'U{row}'] = f'=IF(A{row}="","",IF(I{row}="",0,I{row})+IF(L{row}="",0,L{row})+IF(S{row}="",0,S{row})+IF(T{row}="",0,T{row}))'
        ws[f'U{row}'].alignment = center_align
        ws[f'U{row}'].border = thin_border
        
        # V列: 投資検討（空欄 - 手動入力）
        ws[f'V{row}'].fill = alert_fill if is_portfolio_stock else input_fill
        ws[f'V{row}'].alignment = center_align
        ws[f'V{row}'].border = thin_border
        
        # W列: 投資比率（数式）
        ws[f'W{row}'] = f'=IF(OR(A{row}="",V{row}<>"〇"),"",U{row}/SUMIF($V$6:$V$20,"〇",$U$6:$U$20))'
        ws[f'W{row}'].number_format = '0.0%'
        ws[f'W{row}'].alignment = center_align
        ws[f'W{row}'].border = thin_border
        
        # X列: メモ（空欄 - 手動入力）
        ws[f'X{row}'].fill = alert_fill if is_portfolio_stock else input_fill
        ws[f'X{row}'].alignment = Alignment(horizontal='left', vertical='center')
        ws[f'X{row}'].border = thin_border
        
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
    print(f"更新銘柄数: {len(stock_codes)}銘柄")
    
    if portfolio_alerts:
        print(f"\n⚠️  ポートフォリオ保有中の銘柄（オレンジ色背景）:")
        for code in portfolio_alerts:
            print(f"   - {code}")
        print(f"\n注意: これらの銘柄はポートフォリオに残っています。")
        print(f"      売却済みの場合はポートフォリオシートから削除してください。")
    
    print("\n✅ スクリーニングシート更新完了!")

def get_screening_criteria():
    """
    スクリーニング条件を取得
    
    Returns:
        dict: スクリーニング条件
    """
    print("\n" + "=" * 60)
    print("📊 スクリーニング条件の設定")
    print("=" * 60)
    print("\nスクリーニング条件を入力してください（空Enterでデフォルト値）")
    print()
    
    criteria = {}
    
    # 時価総額
    print("【時価総額】")
    try:
        min_cap = input("  最小時価総額（億円）[デフォルト: 100]: ").strip()
        criteria['min_market_cap'] = float(min_cap) * 100000000 if min_cap else 10000000000
    except:
        criteria['min_market_cap'] = 10000000000  # 100億円
    
    # PER
    print("\n【PER（株価収益率）】")
    try:
        min_per = input("  最小PER [デフォルト: なし]: ").strip()
        criteria['min_per'] = float(min_per) if min_per else None
        max_per = input("  最大PER [デフォルト: なし]: ").strip()
        criteria['max_per'] = float(max_per) if max_per else None
    except:
        criteria['min_per'] = None
        criteria['max_per'] = None
    
    # PBR
    print("\n【PBR（株価純資産倍率）】")
    try:
        min_pbr = input("  最小PBR [デフォルト: なし]: ").strip()
        criteria['min_pbr'] = float(min_pbr) if min_pbr else None
        max_pbr = input("  最大PBR [デフォルト: なし]: ").strip()
        criteria['max_pbr'] = float(max_pbr) if max_pbr else None
    except:
        criteria['min_pbr'] = None
        criteria['max_pbr'] = None
    
    # ROE
    print("\n【ROE（自己資本利益率）】")
    try:
        min_roe = input("  最小ROE（%）[デフォルト: なし]: ").strip()
        criteria['min_roe'] = float(min_roe) if min_roe else None
    except:
        criteria['min_roe'] = None
    
    # 自己資本比率
    print("\n【自己資本比率】")
    try:
        min_equity = input("  最小自己資本比率（%）[デフォルト: なし]: ").strip()
        criteria['min_equity_ratio'] = float(min_equity) if min_equity else None
    except:
        criteria['min_equity_ratio'] = None
    
    # 売買代金
    print("\n【売買代金】")
    try:
        min_value = input("  最小売買代金（億円）[デフォルト: なし]: ").strip()
        criteria['min_trading_value'] = float(min_value) if min_value else None
    except:
        criteria['min_trading_value'] = None
    
    return criteria

def check_screening_criteria(data, criteria):
    """
    銘柄がスクリーニング条件を満たすかチェック
    
    Args:
        data: 銘柄データ
        criteria: スクリーニング条件
    
    Returns:
        bool: 条件を満たす場合True
    """
    # 時価総額（必須条件）
    if data['market_cap'] is None or data['market_cap'] < criteria['min_market_cap']:
        return False
    
    # PER（最小値）
    if criteria['min_per'] is not None:
        if data['trailing_pe'] is None or data['trailing_pe'] < criteria['min_per']:
            return False
    
    # PER（最大値）
    if criteria['max_per'] is not None:
        if data['trailing_pe'] is None or data['trailing_pe'] > criteria['max_per']:
            return False
    
    # PBR（最小値）
    if criteria['min_pbr'] is not None:
        if data['price_to_book'] is None or data['price_to_book'] < criteria['min_pbr']:
            return False
    
    # PBR（最大値）
    if criteria['max_pbr'] is not None:
        if data['price_to_book'] is None or data['price_to_book'] > criteria['max_pbr']:
            return False
    
    # ROE
    if criteria['min_roe'] is not None:
        if data['return_on_equity'] is None or data['return_on_equity'] < criteria['min_roe']:
            return False
    
    # 自己資本比率
    if criteria['min_equity_ratio'] is not None:
        if data['equity_ratio'] is None or data['equity_ratio'] < criteria['min_equity_ratio']:
            return False
    
    # 売買代金
    if criteria['min_trading_value'] is not None:
        if data['trading_value'] is None or data['trading_value'] < criteria['min_trading_value']:
            return False
    
    return True

def auto_screening(max_stocks=15):
    """
    自動スクリーニング：条件に合う銘柄を検索
    
    Args:
        max_stocks: 最大取得銘柄数
    
    Returns:
        list: 条件に合う銘柄コードのリスト
    """
    print("\n🔍 自動スクリーニングを開始します...")
    print("=" * 60)
    
    # 日本の主要銘柄リスト（例）
    # 実際にはもっと多くの銘柄を対象にできます
    candidate_codes = [
        # プライム市場の主要銘柄
        '7203', '6758', '6920', '4063', '8035', '9984', '6861', '6501',
        '7974', '4502', '4503', '8306', '8316', '7751', '6971', '6702',
        '4519', '4568', '6954', '6981', '4324', '9433', '2914', '4911',
        '6367', '7267', '4452', '4523', '6178', '3382', '4704', '9697',
        '6098', '2801', '8058', '8031', '3861', '4661', '6952', '7269',
        '6976', '6645', '4188', '4901', '7733', '6273', '6479', '7832',
        '4543', '6503', '7201', '7270', '9020', '9021', '4755', '6273'
    ]
    
    matched_stocks = []
    criteria = get_screening_criteria()
    
    print("\n" + "=" * 60)
    print("🔍 スクリーニング実行中...")
    print("=" * 60)
    print(f"対象銘柄数: {len(candidate_codes)}銘柄")
    print(f"最大取得数: {max_stocks}銘柄")
    print()
    
    for idx, code in enumerate(candidate_codes, 1):
        if len(matched_stocks) >= max_stocks:
            break
        
        print(f"[{idx}/{len(candidate_codes)}] {code} チェック中...", end=" ")
        
        data = get_stock_data(code)
        
        if data is None:
            print("データ取得失敗")
            continue
        
        if check_screening_criteria(data, criteria):
            print("✓ 条件合致！")
            matched_stocks.append(code)
        else:
            print("×")
        
        # API制限を避けるため待機
        time.sleep(0.3)
    
    print("\n" + "=" * 60)
    print(f"✅ スクリーニング完了: {len(matched_stocks)}銘柄が条件に合致")
    print("=" * 60)
    
    if matched_stocks:
        print("\n【合致した銘柄】")
        for code in matched_stocks:
            print(f"  - {code}")
    
    return matched_stocks

def main():
    """メイン処理"""
    print("=" * 60)
    print("📊 投資管理テンプレート - スクリーニングシート自動更新")
    print("=" * 60)
    
    filepath = None
    
    # コマンドライン引数がある場合
    if len(sys.argv) >= 2:
        filepath = sys.argv[1]
        print(f"\n📁 指定されたファイル: {filepath}")
    else:
        # GUIファイル選択ダイアログを表示
        if HAS_TKINTER:
            print("\n📁 ファイル選択ダイアログを開きます...")
            print("   （ダイアログが表示されない場合は、タスクバーを確認してください）")
            
            filepath = select_excel_file()
            
            if filepath:
                print(f"✅ 選択されたファイル: {filepath}")
            else:
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
    
    # スクリーニング方法の選択
    print("\n" + "=" * 60)
    print("📝 銘柄の選択方法")
    print("=" * 60)
    print("1. 自動スクリーニング（条件に合う銘柄を自動検索）")
    print("2. 手動入力（銘柄コードを直接入力）")
    print()
    
    stock_codes = []
    
    try:
        choice = input("選択してください (1/2) [デフォルト: 1]: ").strip()
    except EOFError:
        choice = '1'
    
    if not choice:
        choice = '1'
    
    if choice == '1':
        # 自動スクリーニング
        stock_codes = auto_screening(max_stocks=15)
        
        if not stock_codes:
            print("\n❌ 条件に合う銘柄が見つかりませんでした")
            print("条件を緩めるか、手動入力を試してください")
            input("\nEnterキーで終了...")
            sys.exit(1)
    
    elif choice == '2':
        # 手動入力
        print("\n📝 更新する銘柄コードを入力してください")
        print("   （複数の場合はカンマ区切り、例: 7203,6758,6920）")
        print("   空Enter で入力終了")
        print()
        
        while True:
            try:
                user_input = input("銘柄コード: ").strip()
            except EOFError:
                break
            
            if not user_input:
                break
            
            # カンマ区切りで分割
            codes = [code.strip() for code in user_input.split(',')]
            stock_codes.extend(codes)
        
        if not stock_codes:
            print("❌ エラー: 銘柄コードが入力されていません")
            input("\nEnterキーで終了...")
            sys.exit(1)
    
    else:
        print("❌ エラー: 無効な選択です")
        input("\nEnterキーで終了...")
        sys.exit(1)
    
    # 重複を削除
    stock_codes = list(dict.fromkeys(stock_codes))
    
    print(f"\n✅ {len(stock_codes)}銘柄を更新します")
    print(f"   {', '.join(stock_codes)}")
    
    # 確認
    try:
        confirm = input("\n続行しますか？ (y/N): ").strip().lower()
    except EOFError:
        confirm = 'n'
    
    if confirm not in ['y', 'yes']:
        print("キャンセルしました")
        input("\nEnterキーで終了...")
        sys.exit(0)
    
    # スクリーニングシートを更新
    try:
        update_screening_sheet(filepath, stock_codes)
    except Exception as e:
        print(f"\n❌ エラーが発生しました: {str(e)}")
        import traceback
        traceback.print_exc()
    
    # 終了前に待機
    input("\nEnterキーで終了...")

if __name__ == "__main__":
    main()
