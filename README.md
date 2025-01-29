# 金融計算Excel模板生成器 Financial Calculator Template Generator

這個專案可以自動生成一個金融計算用的Excel模板，包含以下功能：
This project can automatically generate a financial calculation Excel template, including the following features:

## 功能列表 Features

- PV（現值）計算 Present Value (PV) calculations
- FV（未來值）計算 Future Value (FV) calculations
- 普通年金現值和終值 Ordinary Annuity calculations
- 成長年金計算 Growing Annuity calculations
- 永續年金計算 Perpetuity calculations
- NPV（淨現值）計算 Net Present Value calculations
- IRR（內部報酬率）計算 Internal Rate of Return calculations
- 股利折現模型（DDM） Dividend Discount Model
- 本益比（P/E）計算 Price-Earnings (P/E) Ratio calculations
- 債券價格計算 Bond Price calculations
- 債券殖利率（YTM）計算 Yield to Maturity (YTM) calculations

## 使用方法 Usage

1. 安裝必要套件：
```bash
pip install -r requirements.txt
```

2. 執行Python腳本：
```bash
python create_finance_template.py
```

3. 檢查生成的Excel檔案：
- 會在當前目錄生成 `金融計算模板.xlsx`
- 打開檔案，直接在對應欄位輸入數值即可使用

## Excel模板使用說明 Template Usage Instructions

1. 每個計算類別都有對應的輸入變數說明
   Each calculation type has corresponding input variable descriptions
2. 在輸入變數欄位中填入相應的數值
   Enter values in the input variable fields
3. 結果會自動計算並顯示在結果欄位中
   Results will be automatically calculated and shown in the result fields

## 注意事項 Notes

- 使用Excel 2010或更新版本
  Use Excel 2010 or newer versions
- 確保Excel中已啟用巨集功能
  Ensure macros are enabled in Excel
- 所有利率輸入請使用小數格式（例如：10% 請輸入 0.1）
  All interest rates should be entered in decimal format (e.g., enter 0.1 for 10%)
