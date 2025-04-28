import datetime
import os
import re
import shutil
import tempfile
import traceback
import zipfile

import pandas as pd
from bs4 import BeautifulSoup
from flask import Flask, jsonify, render_template, request, send_file

app = Flask(__name__)

# 確保上傳目錄存在
UPLOAD_FOLDER = "uploads"
if not os.path.exists(UPLOAD_FOLDER):
    os.makedirs(UPLOAD_FOLDER)

app.config["UPLOAD_FOLDER"] = UPLOAD_FOLDER


def parse_egg_production_table(html_content, encoding="big5"):
    """解析HTML表格內容並只提取特定牧場的數據"""
    try:
        # 使用BeautifulSoup解析HTML
        soup = BeautifulSoup(html_content, "html.parser")

        # 找到表格
        table = soup.find("table", class_="list")
        if not table:
            return None

        # 提取標題行以獲取日期
        header_row = table.find_all("tr")[0]
        header_cells = header_row.find_all("td")
        date_info = ""
        week1 = "0407-0413"  # 默認值
        week2 = "0414-0420"  # 默認值

        if len(header_cells) > 0:
            date_info = header_cells[0].get_text().strip()
            # 提取日期範圍 20250407~20250420 雙周比較
            date_match = re.search(r"(\d{8})~(\d{8})", date_info)
            if date_match:
                date1 = date_match.group(1)
                date2 = date_match.group(2)
                date1_obj = datetime.datetime.strptime(date1, "%Y%m%d")
                date2_obj = datetime.datetime.strptime(date2, "%Y%m%d")
                date1_end = date1_obj + datetime.timedelta(days=6)
                date2_start = date1_end + datetime.timedelta(days=1)
                week1 = f"{date1_obj.strftime('%m%d')}-{date1_end.strftime('%m%d')}"
                week2 = f"{date2_start.strftime('%m%d')}-{date2_obj.strftime('%m%d')}"

        # 需要過濾的特定牧場
        target_farms = [
            "富源畜牧場一場(本A",
            "富源畜牧場一場(本B",
            "富源畜牧場三場(3A)",
            "富源畜牧場三場(3D)",
        ]

        # 提取所有行的數據
        raw_data = []
        for tr in table.find_all("tr")[2:]:  # 跳過標題行
            cells = tr.find_all("td")
            if len(cells) < 30:  # 確保有足夠的儲存格
                continue

            # 獲取牧場名稱
            farm_name = cells[1].get_text().strip()

            # 判斷是否是目標牧場
            is_target_farm = False
            for target in target_farms:
                if target in farm_name:
                    is_target_farm = True
                    break

            # 如果不是目標牧場，跳過該行
            if not is_target_farm:
                continue

            # 基本信息
            row_data = {
                "洗選廠": cells[0].get_text().strip(),
                "牧場": farm_name,
                # 處理雞蛋尺寸分佈 - 3S(42g以下)
                "3S_週1": cells[7].get_text().strip(),
                "3S_週2": cells[8].get_text().strip(),
                # 2S(42g-48g)
                "2S_週1": cells[9].get_text().strip(),
                "2S_週2": cells[10].get_text().strip(),
                # S(48g-54g)
                "S_週1": cells[11].get_text().strip(),
                "S_週2": cells[12].get_text().strip(),
                # M(54g-60g)
                "M_週1": cells[13].get_text().strip(),
                "M_週2": cells[14].get_text().strip(),
                # L(60g-66g)
                "L_週1": cells[15].get_text().strip(),
                "L_週2": cells[16].get_text().strip(),
                # 2L(66g-68g)
                "2L_週1": cells[17].get_text().strip(),
                "2L_週2": cells[18].get_text().strip(),
                # 3L(68g-72g)
                "3L_週1": cells[19].get_text().strip(),
                "3L_週2": cells[20].get_text().strip(),
                # 4L(72g以上)
                "4L_週1": cells[21].get_text().strip(),
                "4L_週2": cells[22].get_text().strip(),
                # 裂紋蛋 E1
                "E1_週1": cells[23].get_text().strip(),
                "E1_週2": cells[24].get_text().strip(),
                # 髒蛋 E2
                "E2_週1": cells[25].get_text().strip(),
                "E2_週2": cells[26].get_text().strip(),
                # 異常蛋 E3
                "E3_週1": cells[27].get_text().strip(),
                "E3_週2": cells[28].get_text().strip(),
                # 破蛋 E4
                "E4_週1": cells[29].get_text().strip(),
                "E4_週2": cells[30].get_text().strip(),
                # 總次級蛋%
                "總次級蛋%_週1": cells[33].get_text().strip(),
                "總次級蛋%_週2": cells[34].get_text().strip(),
            }

            # 設定棟別為牧場名稱中的特定部分
            if "本A" in farm_name:
                row_data["棟別"] = "本A"
            elif "本B" in farm_name:
                row_data["棟別"] = "本B"
            elif "(3A)" in farm_name:
                row_data["棟別"] = "3A"
            elif "(3D)" in farm_name:
                row_data["棟別"] = "3D"
            else:
                row_data["棟別"] = farm_name  # 應該不會執行到這裡，但為了保險起見

            raw_data.append(row_data)

        if not raw_data:
            return pd.DataFrame()

        # 清理數據 - 移除空格並統一格式
        for row in raw_data:
            for key, value in row.items():
                if isinstance(value, str):
                    row[key] = value.replace(" ", "").strip()

        # 創建新的格式化數據
        formatted_data = []

        # 處理每個牧場數據
        for row_data in raw_data:
            try:
                # 清理百分比數據，移除%符號並轉為浮點數
                def clean_percent(value):
                    if isinstance(value, str):
                        return float(value.replace("%", "").strip())
                    return 0.0

                # S+2S<54g (合併3S, 2S, S的數據)
                s_3s_1 = clean_percent(row_data.get("3S_週1", "0"))
                s_2s_1 = clean_percent(row_data.get("2S_週1", "0"))
                s_s_1 = clean_percent(row_data.get("S_週1", "0"))
                s_sum_1 = s_3s_1 + s_2s_1 + s_s_1

                s_3s_2 = clean_percent(row_data.get("3S_週2", "0"))
                s_2s_2 = clean_percent(row_data.get("2S_週2", "0"))
                s_s_2 = clean_percent(row_data.get("S_週2", "0"))
                s_sum_2 = s_3s_2 + s_2s_2 + s_s_2

                # M+L(54-66g) (合併M, L的數據)
                m_1 = clean_percent(row_data.get("M_週1", "0"))
                l_1 = clean_percent(row_data.get("L_週1", "0"))
                ml_sum_1 = m_1 + l_1

                m_2 = clean_percent(row_data.get("M_週2", "0"))
                l_2 = clean_percent(row_data.get("L_週2", "0"))
                ml_sum_2 = m_2 + l_2

                # 2-4L>66g (合併2L, 3L, 4L的數據)
                l2_1 = clean_percent(row_data.get("2L_週1", "0"))
                l3_1 = clean_percent(row_data.get("3L_週1", "0"))
                l4_1 = clean_percent(row_data.get("4L_週1", "0"))
                l_large_sum_1 = l2_1 + l3_1 + l4_1

                l2_2 = clean_percent(row_data.get("2L_週2", "0"))
                l3_2 = clean_percent(row_data.get("3L_週2", "0"))
                l4_2 = clean_percent(row_data.get("4L_週2", "0"))
                l_large_sum_2 = l2_2 + l3_2 + l4_2

                # 聲納+髒(E1+E2)
                e1_1 = clean_percent(row_data.get("E1_週1", "0"))
                e2_1 = clean_percent(row_data.get("E2_週1", "0"))

                e1_2 = clean_percent(row_data.get("E1_週2", "0"))
                e2_2 = clean_percent(row_data.get("E2_週2", "0"))

                # 獲取異常蛋和破蛋數據
                e3_1 = row_data.get("E3_週1", "0").replace("%", "")
                e3_2 = row_data.get("E3_週2", "0").replace("%", "")
                e4_1 = row_data.get("E4_週1", "0").replace("%", "")
                e4_2 = row_data.get("E4_週2", "0").replace("%", "")

                # 總次級蛋%
                total_defect_1 = row_data.get("總次級蛋%_週1", "0").replace("%", "")
                total_defect_2 = row_data.get("總次級蛋%_週2", "0").replace("%", "")

                # 添加第一週的資料
                first_row = {
                    "棟別": row_data["棟別"],
                    "日期": week1,
                    "洗選場": row_data["洗選廠"],
                    "S+2S<54g": f"{s_sum_1:.1f}%",
                    "M+L(54-66g)": f"{ml_sum_1:.1f}%",
                    "2-4L>66g": f"{l_large_sum_1:.1f}%",
                    "裂紋蛋E1": f"{e1_1:.1f}%",
                    "髒蛋E2": f"{e2_1:.1f}%",
                    "異常蛋E3": f"{e3_1}%",
                    "破蛋E4": f"{e4_1}%",
                    "總次級蛋%": f"{total_defect_1}%",
                }
                formatted_data.append(first_row)

                # 添加第二週的資料
                second_row = {
                    "棟別": row_data["棟別"],
                    "日期": week2,
                    "洗選場": row_data["洗選廠"],
                    "S+2S<54g": f"{s_sum_2:.1f}%",
                    "M+L(54-66g)": f"{ml_sum_2:.1f}%",
                    "2-4L>66g": f"{l_large_sum_2:.1f}%",
                    "裂紋蛋E1": f"{e1_2:.1f}%",
                    "髒蛋E2": f"{e2_2:.1f}%",
                    "異常蛋E3": f"{e3_2}%",
                    "破蛋E4": f"{e4_2}%",
                    "總次級蛋%": f"{total_defect_2}%",
                }
                formatted_data.append(second_row)

            except (ValueError, KeyError) as e:
                print(f"處理數據時出錯: {str(e)} for row {row_data}")
                continue

        # 創建DataFrame
        df = pd.DataFrame(formatted_data)

        # 確保所有欄位都存在
        required_columns = [
            "棟別",
            "日期",
            "洗選場",
            "S+2S<54g",
            "M+L(54-66g)",
            "2-4L>66g",
            "裂紋蛋E1",
            "髒蛋E2",
            "異常蛋E3",
            "破蛋E4",
            "總次級蛋%",
        ]
        for col in required_columns:
            if col not in df.columns:
                df[col] = ""

        # 按特定順序排列棟別
        custom_order = {"本A": 0, "本B": 1, "3A": 2}
        df["排序"] = df["棟別"].map(lambda x: custom_order.get(x, 999))
        df = df.sort_values(["排序", "日期"]).drop(columns=["排序"])

        return df
    except Exception as e:
        print(f"解析HTML時出錯: {str(e)}")
        traceback.print_exc()
        return pd.DataFrame()


def process_zip_file(zip_path):
    """處理ZIP檔案並解析內部的HTML檔案"""
    temp_dir = tempfile.mkdtemp()
    try:
        # 解壓縮ZIP檔案
        with zipfile.ZipFile(zip_path, "r") as zip_ref:
            zip_ref.extractall(temp_dir)

        # 尋找所有HTML檔案
        html_files = []
        for root, dirs, files in os.walk(temp_dir):
            for file in files:
                if file.lower().endswith(".html") or file.lower().endswith(".htm"):
                    html_files.append(os.path.join(root, file))

        # 解析每個HTML檔案
        results = []
        for html_file in html_files:
            try:
                # 嘗試不同的編碼
                encodings = ["big5", "utf-8", "gb2312", "gbk"]
                html_content = None

                for encoding in encodings:
                    try:
                        with open(html_file, "r", encoding=encoding) as f:
                            html_content = f.read()
                        break  # 如果成功讀取，跳出循環
                    except UnicodeDecodeError:
                        continue

                if html_content is None:
                    print(f"無法使用已知編碼讀取檔案: {html_file}")
                    continue

                # 解析表格
                df = parse_egg_production_table(html_content)

                if df is not None and not df.empty:
                    # 添加來源文件信息
                    df["來源文件"] = os.path.basename(html_file)
                    results.append(df)
                    print(f"已成功解析 {html_file}")
            except Exception as e:
                print(f"解析 {html_file} 時出錯: {str(e)}")

        # 合併所有結果
        if results:
            combined_df = pd.concat(results, ignore_index=True)
            return combined_df
        else:
            return pd.DataFrame()  # 返回空的DataFrame
    finally:
        # 清理臨時目錄
        shutil.rmtree(temp_dir)


def process_html_file(html_path):
    """處理單個HTML檔案"""
    try:
        # 嘗試不同的編碼
        encodings = ["big5", "utf-8", "gb2312", "gbk"]
        html_content = None

        for encoding in encodings:
            try:
                with open(html_path, "r", encoding=encoding) as f:
                    html_content = f.read()
                break  # 如果成功讀取，跳出循環
            except UnicodeDecodeError:
                continue

        if html_content is None:
            return pd.DataFrame(), "無法使用已知編碼讀取檔案"

        # 解析表格
        df = parse_egg_production_table(html_content)

        if df is not None and not df.empty:
            df["來源文件"] = os.path.basename(html_path)
            return df, "success"
        else:
            return pd.DataFrame(), "未找到表格或表格為空"
    except Exception as e:
        return pd.DataFrame(), f"處理檔案時出錯: {str(e)}"


@app.route("/")
def index():
    return render_template("index.html")


@app.route("/upload", methods=["POST"])
def upload_file():
    if "file" not in request.files:
        return jsonify({"error": "沒有文件"}), 400

    file = request.files["file"]

    if file.filename == "":
        return jsonify({"error": "沒有選擇文件"}), 400

    file_path = os.path.join(app.config["UPLOAD_FOLDER"], file.filename)
    file.save(file_path)

    # 根據文件類型處理
    filename_lower = file.filename.lower()
    if filename_lower.endswith(".zip"):
        df = process_zip_file(file_path)
    elif filename_lower.endswith(".html") or filename_lower.endswith(".htm"):
        df, message = process_html_file(file_path)
        if df.empty:
            return jsonify({"error": message}), 400
    else:
        return jsonify({"error": "不支援的文件格式"}), 400

    if df.empty:
        return jsonify({"error": "沒有找到有效的表格數據"}), 400

    # 將DataFrame轉換為HTML表格 (不顯示索引並去除"來源文件"列)
    if "來源文件" in df.columns:
        df = df.drop(columns=["來源文件"])

    # 將DataFrame轉換為HTML表格
    table_html = df.to_html(classes="table table-striped table-bordered")

    # 創建Excel檔案
    excel_path = os.path.join(app.config["UPLOAD_FOLDER"], "parsed_data.xlsx")
    df.to_excel(excel_path, index=False)

    return jsonify(
        {
            "table_html": table_html,
            "excel_path": "download_excel",
            "csv_path": "download_csv",
        }
    )


@app.route("/download_excel")
def download_excel():
    excel_path = os.path.join(app.config["UPLOAD_FOLDER"], "parsed_data.xlsx")
    return send_file(excel_path, as_attachment=True)


if __name__ == "__main__":
    app.run(debug=True)
