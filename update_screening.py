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

# yfinanceのインストール確認
try:
    import yfinance as yf
except ImportError:
    print("yfinanceがインストールされていません。")
    print("以下のコマンドでインストールしてください:")
    print("  pip install yfinance")
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

def get_stock_data(ticker_code):
    """
    yfinanceで株価・財務データを取得
    
    Args:
        ticker_code: 銘柄コード（例: 7203）
    
    Returns:
        dict: 取得したデータ
    """
    try:
        # 日本株は .T を付ける
        ticker = f"{ticker_code}.T"
        stock = yf.Ticker(ticker)
        info = stock.info
        
        # データを辞書形式で返す
        data = {
            'name': info.get('longName', info.get('shortName', '')),
            'market_cap': info.get('marketCap', None),
            'trailing_pe': info.get('trailingPE', None),
            'price_to_book': info.get('priceToBook', None),
            'return_on_equity': info.get('returnOnEquity', None),
            'revenue_growth': info.get('revenueGrowth', None),
        }
        
        # ROEをパーセント表記に変換
        if data['return_on_equity'] is not None:
            data['return_on_equity'] = data['return_on_equity'] * 100
        
        # 売上成長率をパーセント表記に変換
        if data['revenue_growth'] is not None:
            data['revenue_growth'] = data['revenue_growth'] * 100
        
        return data
        
    except Exception as e:
        print(f"  ⚠️  {ticker_code}: データ取得エラー - {str(e)}")
        return None

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
    
    # データ開始行
    start_row = 6
    current_row = start_row
    
    # 既存データをクリア（6行目以降）
    for row in range(start_row, 21):
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
        ws[f'B{row}'] = data['name'] if data['name'] else ''
        ws[f'B{row}'].fill = alert_fill if is_portfolio_stock else input_fill
        ws[f'B{row}'].alignment = center_align
        ws[f'B{row}'].border = thin_border
        
        # C列: 市場区分（デフォルト: プライム）
        ws[f'C{row}'] = 'プライム'
        ws[f'C{row}'].fill = alert_fill if is_portfolio_stock else input_fill
        ws[f'C{row}'].alignment = center_align
        ws[f'C{row}'].border = thin_border
        
        # D列: 時価総額
        if data['market_cap']:
            ws[f'D{row}'] = data['market_cap'] / 100000000  # 億円単位
            ws[f'D{row}'].number_format = '#,##0'
        ws[f'D{row}'].fill = alert_fill if is_portfolio_stock else input_fill
        ws[f'D{row}'].alignment = center_align
        ws[f'D{row}'].border = thin_border
        
        # E列: 自己資本比率（空欄 - 手動入力）
        ws[f'E{row}'].fill = alert_fill if is_portfolio_stock else input_fill
        ws[f'E{row}'].alignment = center_align
        ws[f'E{row}'].border = thin_border
        
        # F列: 売買代金（空欄 - 手動入力）
        ws[f'F{row}'].fill = alert_fill if is_portfolio_stock else input_fill
        ws[f'F{row}'].alignment = center_align
        ws[f'F{row}'].border = thin_border
        
        # G列: PER
        if data['trailing_pe']:
            ws[f'G{row}'] = data['trailing_pe']
            ws[f'G{row}'].number_format = '0.0'
        ws[f'G{row}'].fill = alert_fill if is_portfolio_stock else input_fill
        ws[f'G{row}'].alignment = center_align
        ws[f'G{row}'].border = thin_border
        
        # H列: PBR
        if data['price_to_book']:
            ws[f'H{row}'] = data['price_to_book']
            ws[f'H{row}'].number_format = '0.0'
        ws[f'H{row}'].fill = alert_fill if is_portfolio_stock else input_fill
        ws[f'H{row}'].alignment = center_align
        ws[f'H{row}'].border = thin_border
        
        # I列: バリュースコア（数式）
        ws[f'I{row}'] = f'=IF(OR(A{row}="",G{row}="",H{row}=""),"",IF(AND(G{row}>=5,G{row}<=10,H{row}>=0.5,H{row}<=0.75),20,IF(AND(G{row}>=5,G{row}<=10,H{row}>0.75,H{row}<=1),18,IF(AND(G{row}>10,G{row}<=15,H{row}>=0.5,H{row}<=0.75),18,IF(AND(G{row}>10,G{row}<=15,H{row}>0.75,H{row}<=1),15,10)))))'
        ws[f'I{row}'].alignment = center_align
        ws[f'I{row}'].border = thin_border
        
        # J列: 売上成長率
        if data['revenue_growth']:
            ws[f'J{row}'] = data['revenue_growth']
            ws[f'J{row}'].number_format = '0.0'
        ws[f'J{row}'].fill = alert_fill if is_portfolio_stock else input_fill
        ws[f'J{row}'].alignment = center_align
        ws[f'J{row}'].border = thin_border
        
        # K列: ROE
        if data['return_on_equity']:
            ws[f'K{row}'] = data['return_on_equity']
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

def main():
    """メイン処理"""
    print("=" * 60)
    print("📊 投資管理テンプレート - スクリーニングシート自動更新")
    print("=" * 60)
    
    # コマンドライン引数のチェック
    if len(sys.argv) < 2:
        print("\n使い方:")
        print("  python update_screening.py <Excelファイルパス>")
        print("\n例:")
        print("  python update_screening.py 投資管理テンプレート.xlsx")
        sys.exit(1)
    
    filepath = sys.argv[1]
    
    # 銘柄コードの入力
    print("\n📝 更新する銘柄コードを入力してください")
    print("   （複数の場合はカンマ区切り、例: 7203,6758,6920）")
    print("   空Enter で入力終了")
    print()
    
    stock_codes = []
    
    while True:
        user_input = input("銘柄コード: ").strip()
        
        if not user_input:
            break
        
        # カンマ区切りで分割
        codes = [code.strip() for code in user_input.split(',')]
        stock_codes.extend(codes)
    
    if not stock_codes:
        print("❌ エラー: 銘柄コードが入力されていません")
        sys.exit(1)
    
    # 重複を削除
    stock_codes = list(dict.fromkeys(stock_codes))
    
    print(f"\n✅ {len(stock_codes)}銘柄を更新します")
    print(f"   {', '.join(stock_codes)}")
    
    # 確認
    confirm = input("\n続行しますか？ (y/N): ").strip().lower()
    if confirm not in ['y', 'yes']:
        print("キャンセルしました")
        sys.exit(0)
    
    # スクリーニングシートを更新
    update_screening_sheet(filepath, stock_codes)

if __name__ == "__main__":
    main()
