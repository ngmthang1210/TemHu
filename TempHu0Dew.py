import pandas as pd
import matplotlib.pyplot as plt
from matplotlib.ticker import MultipleLocator
from matplotlib.backends.backend_pdf import PdfPages
from tkinter import Tk, filedialog, messagebox
import os

def process_file(file_path):
    try:
        df = pd.read_excel(file_path, header=0)
        df.columns = df.columns.str.strip()
        df["Timestamp"] = pd.to_datetime(df["Timestamp"], errors="coerce")
        df = df.dropna(subset=["Timestamp", "Temp.(°C)", "RH(%rh)"])

        serial_number = df["Serial Number"].dropna().iloc[0] if "Serial Number" in df else "UNKNOWN"
        logger_name = df["Logger Name"].dropna().iloc[0] if "Logger Name" in df else "Logger"

        # ==== Tính tick mốc thời gian ====
        n = len(df)
        step = n // 6
        idx_ticks = [0]
        for i in range(1, 6):
            idx_ticks.append(i * step)
        idx_ticks.append(n - 1)
        xtick_times = df.iloc[idx_ticks]["Timestamp"].tolist()

        fig, ax1 = plt.subplots(figsize=(11.7, 8.3))
        fig.patch.set_facecolor('white')
        ax1.plot(df["Timestamp"], df["Temp.(°C)"], color='red', linewidth=1.6)
        if "Temp. LL(°C)" in df:
            ax1.plot(df["Timestamp"], df["Temp. LL(°C)"], color='Red', linestyle='dashed', linewidth=1.32)
        if "Temp. HL(°C)" in df:
            ax1.plot(df["Timestamp"], df["Temp. HL(°C)"], color='red', linestyle='dashed', linewidth=1.35)
        if "Dew Point(°C)" in df:
            z=1 #ax1.plot(df["Timestamp"], df["Dew Point(°C)"], color='navy', linewidth=1.6)
        ax2 = ax1.twinx()
        ax2.plot(df["Timestamp"], df["RH(%rh)"], color='green', linewidth=1.6)
        if "RH HL(%rh)" in df:
            ax2.plot(df["Timestamp"], df["RH HL(%rh)"], color='lime', linestyle='dashed', linewidth=1.35)

        ax1.set_ylabel("Temperature (°C)")
        ax2.set_ylabel("Relative Humidity (%rh)")
        ax1.set_ylim(-30, 60)
        ax2.set_ylim(10, 104)
        ax1.yaxis.set_major_locator(MultipleLocator(9))
        ax1.grid(True, linestyle=':', linewidth=0.5)

        # ==== Đảm bảo tick đầu/cuối đúng hai mép ====
        ax1.set_xlim(df["Timestamp"].iloc[0], df["Timestamp"].iloc[-1])
        ax1.set_xticks(xtick_times)
        ax1.set_xticklabels([dt.strftime('%H:%M:%S\n%d/%m/%Y') for dt in xtick_times], fontsize=10.5)

        # Style tiêu đề, legend, grid, chú thích như bản cũ
        ax1.text(0.5, 1.17, "Temp. and Humi. Logger", transform=ax1.transAxes,
                 fontsize=15, fontweight='bold', ha='center')
        ax1.annotate('', xy=(0.33, 1.15), xytext=(0.67, 1.15,),
                     xycoords='axes fraction',
                     arrowprops=dict(arrowstyle='-', color='black', linewidth=1))
        ax1.text(0.01, 1.11, f"{serial_number} - {logger_name}",
                 transform=ax1.transAxes, fontsize=9, ha='left')

        legend_entries = [
            [("Temp.(°C)", 'red', '-'), ("Temp. LL(°C)", 'red', '--'), ("Temp. HL(°C)", 'red', '--')],
            [("RH(%rh)", 'green', '-'), ("RH HL(%rh)", 'lime', '--'),
             #("Dew Point(°C)", 'navy', '-')
             ]
        ]
        x0, y0, dx, dy = 0.53, 1.11, 0.14, 0.035
        for row_idx, row in enumerate(legend_entries):
            for col_idx, (label, color, style) in enumerate(row):
                xpos = x0 + col_idx * dx
                ypos = y0 - row_idx * dy
                ax1.plot([xpos, xpos + 0.025], [ypos + 0.008, ypos + 0.008],
                         transform=ax1.transAxes, color=color, linestyle=style,
                         linewidth=2, clip_on=False)
                ax1.text(xpos + 0.03, ypos, label, transform=ax1.transAxes,
                         fontsize=9, ha='left', va='bottom', color='black')

        start_time = df["Timestamp"].min().strftime('%d-%b-%y %H:%M:%S')
        end_time = df["Timestamp"].max().strftime('%d-%b-%y %H:%M:%S')
        plt.figtext(0.5, 0.01, f'From: {start_time}  To: {end_time}', ha='center', fontsize=10.5)
        plt.subplots_adjust(left=0.07, right=0.93, top=0.83, bottom=0.18)

        output_path = file_path.replace(".xlsx", "_T.pdf")
        with PdfPages(output_path) as pdf:
            pdf.savefig(fig, dpi=300)
        return os.path.basename(output_path)
    except Exception as e:
        return f"Lỗi: {os.path.basename(file_path)} – {e}"

def run_batch_plot():
    root = Tk()
    root.withdraw()
    file_paths = filedialog.askopenfilenames(
        title="Chọn nhiều file Excel (.xlsx)",
        filetypes=[("Excel files", "*.xlsx")])
    if not file_paths:
        messagebox.showinfo("Huỷ bỏ", "Không có file nào được chọn.")
        return

    log = ""
    for idx, path in enumerate(file_paths):
        result = process_file(path)
        log += f"{idx+1}. {result}\n"

    messagebox.showinfo("Kết quả", f"✅ Xử lý xong {len(file_paths)} file:\n\n{log}")

if __name__ == "__main__":
    run_batch_plot()
