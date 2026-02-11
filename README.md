# 員工請假單生成系統

![version](https://img.shields.io/badge/version-3.0.0-green)
![license](https://img.shields.io/badge/license-Internal%20Use%20Only-blue)
![python](https://img.shields.io/badge/Python-3.7%2B-orange)
![openpyxl](https://img.shields.io/badge/openpyxl-3.1%2B-lightgrey)
![tkinter](https://img.shields.io/badge/tkinter-supported-brightgreen)
![platform](https://img.shields.io/badge/Platform-Windows%20%7C%20macOS%20%7C%20Linux-lightblue)
![icon](icon.png)

---

## 專案說明

**員工請假單生成系統** 是一款基於 Python 開發的圖形化工具，專為企業內部員工請假單自動化生成設計。  
本系統可讀取員工信息庫、支援視覺化日期選擇、自動計算請假時長，並基於 Excel 模板生成標準化請假單文件，全程無需手動編輯 Excel，杜絕格式錯誤與人為疏漏。

---

## 功能特色

- **員工信息自動讀取**：從 `員工名單.xlsx` 加載姓名、工號、部門信息，下拉選擇自動填充。
- **標準化假別選擇**：內建多種帶 `Leave` 後綴的標準假別選項，確保格式統一。
- **日曆化日期選擇**：採用視覺化日曆控件挑選請假日期，避免手動輸入格式錯誤。
- **請假時長自動計算**：選擇結束日期後自動計算天數/小時數（預設每日 8 小時），並校驗日期合法性。
- **申請日期自動填入**：生成請假單時自動在 `H5` 單元格填入當前系統日期。
- **模板復制與重命名**：基於原始模板生成獨立文件，以「工號+請假單」命名，便於歸檔查詢。
- **圖形化操作界面**：基於 TKinter 開發直觀界面，無代碼基礎也可輕鬆操作。

---

## 環境依賴

### 必要套件
| 套件名稱 | 作用 | 安裝指令 |
|----------|------|----------|
| `openpyxl` | 讀寫 Excel (.xlsx) 文件 | `pip install openpyxl` |
| `tkinter` | 構建圖形化界面（Python 自帶） | - |
| `tkcalendar` | 提供日曆選擇控件 | `pip install tkcalendar` |
| `PyInstaller`（選用） | 打包為 EXE 可執行文件 | `pip install pyinstaller` |

### 支援系統
- **Windows**：7/10/11（推薦，打包後 EXE 可直接運行）
- **macOS/Linux**：需安裝 Python 環境，界面顯示可能略有差異

---

## 安裝指南

1. **準備前置文件**
   - 編輯 `員工名單.xlsx`：必須包含「部門」「姓名」「工號」列，填入員工信息並保存為 `.xlsx` 格式。
   - 準備 `請假單 Leave App Form.xlsx`：配置好固定格式，確保 `H5`、`B6`、`E6`、`I6`、`B8`、`B12`、`B24`、`I24` 單元格可編輯。
   - （選用）準備 `my_icon.ico`：用於 EXE 打包的自定義圖標（建議 256x256 尺寸）。

2. **安裝相依套件**
   ```bash
   pip install openpyxl tkcalendar
   ```

3. **（選用）打包為 EXE 文件**
   ```bash
   pyinstaller -F -w -i my_icon.ico --name "員工請假單生成系統" leave_application.py
   ```

---

## 啟動方式

### 方式 1：直接運行 Python 腳本
```bash
# 切換到腳本目錄
cd C:\你的專案路徑\員工請假單生成系統

# 執行主程式
python leave_application.py
```

### 方式 2：運行打包後的 EXE 文件
1. 將 `config.txt`、`員工名單.xlsx`、`請假單 Leave App Form.xlsx` 複製到 EXE 同級目錄。
2. 雙擊 `員工請假單生成系統.exe` 啟動（無需安裝 Python 環境）。

---

## 操作流程

1. **選擇員工**：從下拉框挑選員工，部門和工號自動填充。
2. **選擇假別**：單選對應的請假類型（如 `Personal Leave`、`Sick Leave` 等）。
3. **填寫請假說明**：在文本框輸入請假原因/備註。
4. **選擇請假日期**：
   - 點擊「請假開始日期」日曆控件選擇起始日期。
   - 點擊「請假結束日期」日曆控件選擇結束日期（自動計算時長）。
   - 若結束日期早於開始日期，系統自動修正並給出提示。
5. **生成請假單**：點擊按鈕，在同目錄生成「工號+請假單 Leave App Form.xlsx」文件。

---

## 檔案結構

```
├── 員工請假單生成系統/
│   ├── leave_application.py       # 主程式腳本
│   ├── my_icon.ico                # EXE 打包自定義圖標（選用）
│   ├── config.txt                 # 配置文件（Excel 路徑）
│   ├── 員工名單.xlsx              # 員工信息數據源（必填）
│   ├── 請假單 Leave App Form.xlsx # 請假單模板（必填）
│   └── README.md                  # 說明文檔
```

---

## 常見問題排查

| 問題現象 | 解決方案 |
|----------|----------|
| 找不到 openpyxl/tkcalendar 模塊 | 重新安裝套件：<br>`pip uninstall openpyxl tkcalendar -y`<br>`pip install openpyxl tkcalendar` |
| EXE 運行閃退無提示 | 去掉 `-w` 參數重新打包，查看控制台報錯（通常是文件缺失/路徑錯誤） |
| 日曆控件無法顯示 | 更新 tkcalendar：`pip install --upgrade tkcalendar` |
| 生成的請假單部分單元格為空 | 1. 確認員工名單列名正確<br>2. 確認模板單元格未被保護<br>3. 確認已選擇員工/假別 |
| 自定義圖標不生效 | 1. 確認格式為 .ico 且尺寸 256x256<br>2. 確認打包指令參數正確<br>3. 刷新文件夾緩存 |

---

## 自定義配置

- **修改請假時長計算規則**：編輯 `calculate_total_hours_auto` 方法，調整 `total_hours = delta_days * 8` 中的「8」為實際每日工作時長。
- **新增假別選項**：在 `build_ui` 方法的 `leave_types` 列表中，按「假別名稱+ Leave」格式添加選項。

---

## 注意事項

- 所有核心文件（`config.txt`、`員工名單.xlsx`、`請假單模板`）必須與腳本/EXE 同目錄。
- 生成請假單前請關閉原始模板文件，避免文件被占用。
- 同名請假單文件會被覆蓋，請及時歸檔重要文件。
- 請確保員工信息數據的保密性，避免敏感信息洩露。

---

## 版本更新

- **初始版本**：實現核心生成功能，支援手動輸入日期和計算時長。
- **v2.0**：新增日曆選擇控件，實現請假時長自動計算。
- **v2.5**：新增申請日期自動填入，校驗日期合法性，統一假別格式。
- **v3.0**：優化模板複製邏輯，支援工號命名文件，提升歸檔便捷性。

---

## 免責聲明

- 本系統僅供企業內部使用，不適用於商業用途。
- 使用者需自行確保員工信息的保密性，開發者不承擔信息洩露責任。
- 因模板格式錯誤、環境缺失或操作不當導致的數據丟失，開發者不承擔責任。

---

> 本專案僅供企業內部管理使用，如需商業化部署請自行評估法律與合規風險。

### 總結
1. 本系統基於 Python 實現了視覺化的員工請假單自動生成功能，核心依賴 openpyxl 處理 Excel、tkcalendar 實現日期選擇。
2. 部署時需確保員工名單和請假單模板文件與程式同目錄，執行前需安裝對應依賴套件。
3. 支援直接運行 Python 腳本或打包為 EXE 文件，操作流程簡單，並提供了常見問題的解決方案。
