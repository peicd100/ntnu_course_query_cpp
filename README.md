# 師大課程查詢系統

## 1. 專案概述

本專案為桌面 GUI 課程查詢與排課工具。使用者將課程 Excel 檔放入 `user_data/course_inputs/`，程式會讀取資料並輸出排課結果至 `user_data/user_schedules/`。

## 2. 重要變數

- ENV_NAME：`ntnu_course_query_cpp`（英文小寫+底線；亦作為 GitHub repository 命名依據）
- EXE_NAME：`師大課程查詢系統`（workspace root 資料夾名稱 basename；亦作為 .exe 命名依據）

## 3. workspace root 定義

- workspace root（絕對路徑）：`D:\x1064\PEICD100\0_python\1_師大課程查詢系統\1_師大課程查詢系統_測試版\師大課程查詢系統`
- EXE_NAME 為上述資料夾名稱 basename。
- README 內所有相對路徑均以 workspace root 為準。

## 4. 檔案與資料夾結構（樹狀；最小必要集合）

```
師大課程查詢系統/
├─ README.md
├─ app_main.py
├─ app_constants.py
├─ app_excel.py
├─ app_mainwindow.py
├─ app_timetable_logic.py
├─ app_user_data.py
├─ app_utils.py
├─ app_widgets.py
├─ app_workers.py
├─ user_data/
│  ├─ course_inputs/
│  │  └─ *.xls / *.xlsx
│  └─ user_schedules/
│     └─ <username>/
│        ├─ history/
│        │  └─ <YYYYMMDD_HHMMSS>.xlsx
│        └─ best_schedule/
│           ├─ <最佳課表>_第<序號>.xlsx
│           └─ best_schedule_cache.json
├─ dist/
└─ build/
```

- 專案輸入檔案存放位置：`user_data/course_inputs/`
- 專案輸出檔案存放位置：`user_data/user_schedules/<username>/history/`、`user_data/user_schedules/<username>/best_schedule/`

## 5. Python 檔名規則（app_main.py + app_*.py 同層）

- 入口檔：`app_main.py`
- 其餘模組：`app_*.py`，與 `app_main.py` 同層
- 不新增其他入口檔

## 6. user_data/ 規範

- 所有輸入/輸出/設定預設放在 `user_data/`
- 課程輸入：`user_data/course_inputs/`（`.xls` / `.xlsx`）
- 排課輸出：`user_data/user_schedules/<username>/history/`
- 最佳課表：`user_data/user_schedules/<username>/best_schedule/`
- 執行 `.exe` 時，`user_data/` 需與 `.exe` 同層放置

## 7. Conda 環境（ENV_NAME）規範

- 僅使用 `ntnu_course_query_cpp` 作為專案環境名稱
- 禁止對 base 環境 install/remove
- 套件安裝以 conda 為主；只有在 conda 不可行時才使用 pip
- pip 必須在已啟用 `ntnu_course_query_cpp` 的情況下執行

## 8. 從零開始安裝流程（可一鍵複製；註解不得同列命令）

目標 shell：Windows CMD。

A. 推薦方案（conda 優先）

```
conda create -n ntnu_course_query_cpp python=3.10 -y
call "C:\ProgramData\Anaconda3\Scripts\activate.bat" "C:\ProgramData\Anaconda3" && conda activate base && conda activate ntnu_course_query_cpp
conda install -c conda-forge pyside6 numpy pandas openpyxl xlrd pyinstaller -y
python -c "import PySide6, numpy, pandas, openpyxl, xlrd; import app_main; print('OK')" && python app_main.py
```

B. 備援方案（conda + pip）

若 conda 的 PySide6 在本機無法取得或版本不相容，改以 pip 安裝 PySide6。

```
conda create -n ntnu_course_query_cpp python=3.10 -y
call "C:\ProgramData\Anaconda3\Scripts\activate.bat" "C:\ProgramData\Anaconda3" && conda activate base && conda activate ntnu_course_query_cpp
conda install -c conda-forge numpy pandas openpyxl xlrd pyinstaller -y
pip install PySide6
python -c "import PySide6, numpy, pandas, openpyxl, xlrd; import app_main; print('OK')" && python app_main.py
```

C. 最後手段（pip only）

```
conda create -n ntnu_course_query_cpp python=3.10 -y
call "C:\ProgramData\Anaconda3\Scripts\activate.bat" "C:\ProgramData\Anaconda3" && conda activate base && conda activate ntnu_course_query_cpp
pip install PySide6 numpy pandas openpyxl xlrd pyinstaller
python -c "import PySide6, numpy, pandas, openpyxl, xlrd; import app_main; print('OK')" && python app_main.py
```

建議使用 A 方案，因為相容性與安裝成功率最佳，且本專案無需 GPU 版套件。

## 9. 測試方式

基本啟動測試（GUI 會短暫啟動後自動關閉）：

```
python app_main.py --smoke-test
```

手動測試（正常啟動 GUI）：

```
python app_main.py
```

## 10. 打包成 .exe（可複製指令）

在 `ntnu_course_query_cpp` 環境內執行：

```
call "C:\ProgramData\Anaconda3\Scripts\activate.bat" "C:\ProgramData\Anaconda3" && conda activate base && conda activate ntnu_course_query_cpp
pyinstaller --noconsole --onedir --name "師大課程查詢系統" app_main.py -y
```

輸出位置：

- `dist/師大課程查詢系統/師大課程查詢系統.exe`

若要在未安裝 Python/Conda 的環境使用 `.exe`，請將 `user_data/` 放在 `.exe` 同層。

## 11. 使用者要求

- README 必須維持本文件之章節規範與凍結區塊
- 所有輸入/輸出/設定預設放在 `user_data/`
- 禁止修改系統設定或全域安裝
- Git 操作僅提供指令範例，實際操作由使用者執行

## 12. GitHub操作指令

# 初始化
```
(
echo.
echo # ignore build outputs
echo dist/
echo build/
)>> .gitignore
git init
git branch -M main
git remote add origin https://github.com/peicd100/ntnu_course_query_cpp.git
git add .
git commit -m "PEICD100"
git push -u origin main
```

# 例行上傳
```
git add .
git commit -m "PEICD100"
git push -u origin main
```

# 還原成Git Hub最新資料
```
git rebase --abort || echo "No rebase in progress" && git fetch origin && git switch main && git reset --hard origin/main && git clean -fd && git status
```

# 查看儲存庫
```
git remote -v
```

# 克隆儲存庫
```
git clone https://github.com/peicd100/ntnu_course_query_cpp.git
```
