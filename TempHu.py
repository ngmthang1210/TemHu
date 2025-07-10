import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.dates as mdates

# Đọc lại dữ liệu (nếu chưa có df)
df = pd.read_excel("Từ 21.04 đến 25.04.25 TB.CB.01.xls.xlsx")
df.columns = [str(c).strip() for c in df.columns]

# Chuyển cột thời gian thành kiểu ngày giờ
df['Timestamp'] = pd.to_datetime(df['Timestamp'])

# Tạo figure và trục
fig, ax1 = plt.subplots(figsize=(14, 7))

# Đường nhiệt độ chính (màu đỏ)
ax1.plot(df['Timestamp'], df['Temp.(°C)'], color='red', label='Temp.(°C)', linewidth=2)
# Nét đứt đỏ: giới hạn LL và HL
ax1.plot(df['Timestamp'], df['Temp. LL(°C)'], color='red', linestyle='dashed', label='Temp. LL(°C)')
ax1.plot(df['Timestamp'], df['Temp. HL(°C)'], color='red', linestyle='dashed', label='Temp. HL(°C)')
# Đường Dew Point (xanh dương)
ax1.plot(df['Timestamp'], df['Dew Point(°C)'], color='blue', label='Dew Point(°C)')

ax1.set_xlabel('Time')
ax1.set_ylabel('Temperature (°C)', color='red')
ax1.tick_params(axis='y', labelcolor='red')
ax1.grid(True, which='both', axis='both', linestyle='--', alpha=0.3)

# Format trục X cho đẹp
ax1.xaxis.set_major_formatter(mdates.DateFormatter('%d/%m %H:%M'))
plt.xticks(rotation=30)

# Vẽ RH trên trục phụ (ax2)
ax2 = ax1.twinx()
ax2.plot(df['Timestamp'], df['RH(%rh)'], color='limegreen', label='RH(%rh)', linewidth=2)
ax2.plot(df['Timestamp'], df['RH HL(%rh)'], color='limegreen', linestyle='dashed', label='RH HL(%rh)')

ax2.set_ylabel('Humidity (%RH)', color='green')
ax2.tick_params(axis='y', labelcolor='green')

# Ghép chú thích từ cả 2 trục
lines, labels = ax1.get_legend_handles_labels()
lines2, labels2 = ax2.get_legend_handles_labels()
plt.legend(lines + lines2, labels + labels2, loc='upper left', fontsize=10)

# Tiêu đề biểu đồ
plt.title('Temp. and Humi. Logger\nTừ 21.04 đến 25.04.25 TB.CB.01', fontsize=14, fontweight='bold')

plt.tight_layout()
plt.show()
