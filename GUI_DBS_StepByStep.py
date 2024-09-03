import os
import sys
import pandas as pd
import matplotlib.pyplot as plt
import numpy as np
from tkinter import Tk, filedialog, Button, Label, Radiobutton, StringVar
from PIL import Image, ImageTk
from openpyxl import Workbook
from openpyxl.drawing.image import Image as ExcelImage
from openpyxl.utils.dataframe import dataframe_to_rows


def resource_path(relative_path):
    """ Lấy đường dẫn tuyệt đối tới tài nguyên, hoạt động cho cả khi phát triển và khi đã đóng gói bằng PyInstaller """
    try:
        # PyInstaller tạo ra một thư mục tạm và lưu trữ đường dẫn thực thi trong _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)


def analyze_files_dbscan():
    """ Hàm phân tích cho tùy chọn DBSCAN """
    # Hiển thị hộp thoại chọn thư mục
    folder_path = filedialog.askdirectory()
    if not folder_path:
        return

    # Tạo danh sách để lưu trữ kết quả
    results = []

    # Tạo một workbook mới để lưu dữ liệu và hình ảnh
    wb = Workbook()
    ws_summary = wb.active
    ws_summary.title = "Summary"

    # Duyệt qua tất cả các tệp trong thư mục
    for file_name in os.listdir(folder_path):
        if file_name.endswith('.phsp'):  # Chỉ xử lý các tệp có đuôi .phsp
            file_path = os.path.join(folder_path, file_name)

            # Đọc tệp vào DataFrame
            columns = ['Event_number', 'Single_strand_breaks', 'Double_strand_breaks',
                       'Complex_strand_breaks', 'Cluster_sizes', 'Cluster_size_weights']
            data = pd.read_csv(file_path, delim_whitespace=True, header=None, names=columns)

            # Tính tổng số strand breaks cho tệp hiện tại
            sum_single_strand_breaks = data['Single_strand_breaks'].sum()
            sum_double_strand_breaks = data['Double_strand_breaks'].sum()
            sum_complex_strand_breaks = data['Complex_strand_breaks'].sum()

            # Lưu kết quả vào một từ điển
            file_result = {
                'File Name': file_name,
                'Single Strand Breaks': sum_single_strand_breaks,
                'Double Strand Breaks': sum_double_strand_breaks,
                'Complex Strand Breaks': sum_complex_strand_breaks
            }

            # Thêm từ điển vào danh sách kết quả
            results.append(file_result)

            # Vẽ biểu đồ cho từng tệp
            plt.figure(figsize=(8, 6))
            plt.bar(['Single Strand Breaks', 'Double Strand Breaks', 'Complex Strand Breaks'],
                    [sum_single_strand_breaks, sum_double_strand_breaks, sum_complex_strand_breaks])
            plt.xlabel('Strand Break Type')
            plt.ylabel('Sum')
            plt.title(f'Sum of Strand Breaks for {file_name}')
            plt.grid(axis='y', linestyle='--', alpha=0.7)

            # Lưu biểu đồ thành hình ảnh
            chart_path = os.path.join(folder_path, f"{file_name}.png")
            plt.savefig(chart_path)
            plt.close()

            # Thêm một sheet mới vào workbook cho mỗi file
            ws = wb.create_sheet(title=file_name[:30])  # Tên sheet phải ngắn hơn 31 ký tự
            img = ExcelImage(chart_path)
            ws.add_image(img, 'A1')

    # Chuyển đổi kết quả thành DataFrame và ghi vào trang tổng hợp
    results_df = pd.DataFrame(results)
    for r in dataframe_to_rows(results_df, index=False, header=True):
        ws_summary.append(r)

    # Hiển thị hộp thoại để chọn đường dẫn lưu tệp
    output_path = filedialog.asksaveasfilename(defaultextension=".xlsx",
                                               filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")])

    # Kiểm tra nếu người dùng đã chọn một đường dẫn
    if output_path:
        # Lưu workbook vào file Excel
        wb.save(output_path)

        # Thông báo thành công
        print(results_df)
        print(f"Excel file saved to: {output_path}")
        label.config(text=f"Result saved to: {output_path}")
    else:
        label.config(text="Save operation was canceled.")


def analyze_files_step_by_step():
    """ Hàm phân tích cho tùy chọn Step-by-Step """
    # Hiển thị hộp thoại chọn thư mục
    folder_path = filedialog.askdirectory()
    if not folder_path:
        return

    # Tạo danh sách để lưu trữ kết quả
    results = []

    # Duyệt qua tất cả các tệp trong thư mục
    for file_name in os.listdir(folder_path):
        if file_name.endswith('.phsp'):  # Chỉ xử lý các tệp có đuôi .phsp
            file_path = os.path.join(folder_path, file_name)

            # Đọc tệp vào DataFrame
            columns = [
                "Energy_imparted_per_event", "Dose_per_event_Gy", "DSB/Gy/Gbp", "SSB/Gy/Gbp",
                "SB/Gy/Gbp", "SSB+/Gy/Gbp", "DSB+/Gy/Gbp", "MoreComplexDamage/Gy/Gbp",
                "BD/Gy/Gbp", "DSBs", "DSBs_Direct", "DSBs_Indirect", "DSBs_Hybrid",
                "DSBs_Direct_WithOneQuasiDirect", "DSBs_Direct_WithBothQuasiDirect",
                "DSBs_Hybrid_WithOneQuasiDirect", "SSBs", "SSBs_Direct", "SSBs_QuasiDirect",
                "SSBs_Indirect", "SBs", "SBs_Direct", "SBs_QuasiDirect", "SBs_Indirect",
                "SSB+s", "DSB+s", "More_complex_damages", "BDs", "BDs_Direct", "BDs_QuasiDirect",
                "BDs_Indirect", "Foci_150nm", "Foci_500nm"
            ]
            df = pd.read_csv(file_path, delim_whitespace=True, names=columns)

            # Chọn các cột mà bạn muốn so sánh
            selected_columns = [
                "Dose_per_event_Gy", "DSBs", "DSBs_Direct", "DSBs_Indirect", "DSBs_Hybrid",
                "SSBs", "SSBs_Direct", "SSBs_QuasiDirect", "SSBs_Indirect", "DSB/Gy/Gbp", "SSB/Gy/Gbp"
            ]
            df_selected = df[selected_columns]

            total_dose = df_selected["Dose_per_event_Gy"].sum()
            total_DSBs = df_selected["DSBs"].sum()
            total_SSBs = df_selected["SSBs"].sum()

            total_DSBs_Direct = df_selected["DSBs_Direct"].sum()
            total_DSBs_Indirect = df_selected["DSBs_Indirect"].sum()
            total_DSBs_Hybrid = df_selected["DSBs_Hybrid"].sum()

            total_SSBs_Direct = df_selected["SSBs_Direct"].sum()
            total_SSBs_QuasiDirect = df_selected["SSBs_QuasiDirect"].sum()
            total_SSBs_Indirect = df_selected["SSBs_Indirect"].sum()

            ssb_dsb_ratio = total_SSBs / total_DSBs if total_DSBs != 0 else 0

            # Tính phần trăm cho các thành phần của DSBs và SSBs
            percentage_DSBs_Direct = (total_DSBs_Direct / total_DSBs) * 100 if total_DSBs != 0 else 0
            percentage_DSBs_Indirect = (total_DSBs_Indirect / total_DSBs) * 100 if total_DSBs != 0 else 0
            percentage_DSBs_Hybrid = (total_DSBs_Hybrid / total_DSBs) * 100 if total_DSBs != 0 else 0

            percentage_SSBs_Direct = (total_SSBs_Direct / total_SSBs) * 100 if total_SSBs != 0 else 0
            percentage_SSBs_QuasiDirect = (total_SSBs_QuasiDirect / total_SSBs) * 100 if total_SSBs != 0 else 0
            percentage_SSBs_Indirect = (total_SSBs_Indirect / total_SSBs) * 100 if total_SSBs != 0 else 0

            # Thêm kết quả vào bảng so sánh
            stats = {
                'File': file_name,
                'total_dose': total_dose,
                'Number_DSBs': total_DSBs,
                'Number_DSBs_Direct': total_DSBs_Direct,
                'Number_DSBs_Indirect': total_DSBs_Indirect,
                'Number_DSBs_Hybrid': total_DSBs_Hybrid,
                'Number_SSBs': total_SSBs,
                'Number_SSBs_Direct': total_SSBs_Direct,
                'Number_SSBs_QuasiDirect': total_SSBs_QuasiDirect,
                'Number_SSBs_Indirect': total_SSBs_Indirect,
                'Percentage_DSBs_Direct': percentage_DSBs_Direct,
                'Percentage_DSBs_Indirect': percentage_DSBs_Indirect,
                'Percentage_DSBs_Hybrid': percentage_DSBs_Hybrid,
                'Percentage_SSBs_Direct': percentage_SSBs_Direct,
                'Percentage_SSBs_QuasiDirect': percentage_SSBs_QuasiDirect,
                'Percentage_SSBs_Indirect': percentage_SSBs_Indirect,
                'DSB/Gy/Gbp': df_selected["DSB/Gy/Gbp"].sum(),
                'SSB/Gy/Gbp': df_selected["SSB/Gy/Gbp"].sum(),
                'SSB/DSBs Ratio': ssb_dsb_ratio
            }

            results.append(stats)

            # Vẽ biểu đồ hình tròn nếu các giá trị hợp lệ
            sizes_DSBs = [percentage_DSBs_Direct, percentage_DSBs_Indirect, percentage_DSBs_Hybrid]
            sizes_SSBs = [percentage_SSBs_Direct, percentage_SSBs_QuasiDirect, percentage_SSBs_Indirect]

            if not any(np.isnan(sizes_DSBs)) and sum(sizes_DSBs) > 0 and not any(np.isnan(sizes_SSBs)) and sum(
                    sizes_SSBs) > 0:
                plt.figure(figsize=(14, 7))

                # Biểu đồ hình tròn cho DSBs (bên trái)
                plt.subplot(1, 2, 1)
                labels_DSBs = ['Direct', 'Indirect', 'Hybrid']
                plt.pie(sizes_DSBs, labels=labels_DSBs, autopct='%1.1f%%', startangle=140)
                plt.title(f"DSBs for {os.path.basename(file_path)}")

                # Biểu đồ hình tròn cho SSBs (bên phải)
                plt.subplot(1, 2, 2)
                labels_SSBs = ['Direct', 'QuasiDirect', 'Indirect']
                plt.pie(sizes_SSBs, labels=labels_SSBs, autopct='%1.1f%%', startangle=140)
                plt.title(f"SSBs for {os.path.basename(file_path)}")

                # Hiển thị biểu đồ
                plt.tight_layout()
                plt.savefig(os.path.join(folder_path,
                                         f"DSBs_SSBs_PieChart_{os.path.basename(file_path).replace('.phsp', '')}.png"))
                plt.close()

    # Tạo DataFrame từ bảng so sánh
    comparison_df = pd.DataFrame(results)

    # Hiển thị hộp thoại để chọn đường dẫn lưu tệp Excel ngay lập tức sau khi xử lý xong
    output_path = filedialog.asksaveasfilename(defaultextension=".xlsx",
                                               filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")])

    if output_path:
        # Tạo file Excel từ DataFrame
        comparison_df.to_excel(output_path, index=False)
        print(f"Excel file saved to: {output_path}")
        label.config(text=f"Result saved to: {output_path}")
    else:
        label.config(text="Save operation was canceled.")


def on_analyze_button_click():
    """ Hàm được gọi khi người dùng nhấn nút phân tích """
    selected_method = method_var.get()
    if selected_method == "DBSCAN":
        analyze_files_dbscan()
    elif selected_method == "Step-by-Step":
        analyze_files_step_by_step()


# Tạo cửa sổ chính của Tkinter
root = Tk()
root.title("Analyze Phase Space Files")

# Thiết lập kích thước cố định cho cửa sổ
root.geometry("600x500")
root.resizable(False, False)

# Thêm nhãn chào mừng
welcome_label = Label(root, text="Welcome to Lee Lab", font=("Arial", 14, "bold"), fg="blue", anchor='center',
                      justify='center')
welcome_label.pack(pady=10)

# Đường dẫn đến logo (sử dụng hàm resource_path)
logo_path = resource_path("banner_branding_880x300.jpg")

# Tải logo từ đường dẫn cục bộ
try:
    logo_image = Image.open(logo_path)
    logo_image = logo_image.resize((350, 90), Image.ANTIALIAS)
    logo_photo = ImageTk.PhotoImage(logo_image)
    logo_label = Label(root, image=logo_photo)
    logo_label.image = logo_photo  # Giữ tham chiếu đến đối tượng ảnh
    logo_label.pack(pady=5)
except Exception as e:
    print(f"Không thể tải hình ảnh: {e}")

# Tạo biến để lưu trữ phương pháp được chọn
method_var = StringVar(value="DBSCAN")

# Thêm nhãn và nút vào giao diện
label = Label(root, text="Select the type of phase space file you would like to analyze", font=("Arial", 14),
              anchor='center', justify='center')
label.pack(pady=10)

# Tạo nút chọn phương pháp (Radio buttons)
radiobutton1 = Radiobutton(root, text="DBSCAN", variable=method_var, value="DBSCAN", font=("Arial", 12))
radiobutton1.pack()

radiobutton2 = Radiobutton(root, text="Step-by-Step", variable=method_var, value="Step-by-Step", font=("Arial", 12))
radiobutton2.pack()

button = Button(root, text="Select and Analyze Files", command=on_analyze_button_click, font=("Arial", 12))
button.pack(pady=20)

# Chạy vòng lặp chính của Tkinter
root.mainloop()

