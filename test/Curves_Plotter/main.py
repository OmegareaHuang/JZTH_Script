import os
import pandas as pd
import matplotlib.pyplot as plt
import argparse


def plot_training_curves(data_folder, y_axis_param, save_path, resolution):
    # 获取数据文件夹中的所有 Excel 文件
    excel_files = [f for f in os.listdir(data_folder) if f.endswith('.xlsx')]

    if not excel_files:
        print(f"No Excel files found in directory: {data_folder}")
        return

    plt.figure(figsize=(10, 6))

    for file in excel_files:
        file_path = os.path.join(data_folder, file)
        print(f"Reading file: {file_path}")

        try:
            df = pd.read_excel(file_path, engine='openpyxl')  # 指定引擎为 openpyxl
        except Exception as e:
            print(f"Error reading {file_path}: {e}")
            continue

        epoch = df['epoch_name'].str.extract(r'(\d+)').astype(int).squeeze()
        y_data = df[y_axis_param]

        # 绘制曲线
        plt.plot(epoch, y_data, label=os.path.splitext(file)[0])

        # 在曲线末端标注文件名
        plt.text(epoch.iloc[-1], y_data.iloc[-1], os.path.splitext(file)[0], fontsize=8, verticalalignment='bottom')

    plt.xlabel('Epoch')
    plt.ylabel(y_axis_param)
    plt.title(f'Training Curves for {y_axis_param}')
    plt.legend()
    plt.grid(True)

    # 修改 x 轴刻度，使其每 10 个 epoch 显示一次
    plt.xticks(ticks=range(0, len(epoch), 10), labels=epoch[::10], rotation=45)

    # 修改 save_path，拼接 Column 参数的值
    base, ext = os.path.splitext(save_path)
    save_path = os.path.join(os.path.dirname(base), f"{y_axis_param}_{os.path.basename(base)}{ext}")

    # 确保文件不被覆盖，添加后缀区分
    i = 1
    original_save_path = save_path
    while os.path.exists(save_path):
        save_path = f"{os.path.splitext(original_save_path)[0]}_{i}{ext}"
        i += 1

    # 保存图像，设置分辨率
    plt.savefig(save_path, dpi=resolution * 100)
    plt.show()


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Plot training curves from multiple Excel files")
    parser.add_argument('--data', type=str, default='excel', help='Path to the folder containing the Excel files')
    parser.add_argument('--Column', type=str, required=True, help='Column name to be plotted on the Y axis')
    parser.add_argument('--save', type=str, default='IMG/curves.png', help='Path to save the plot image')
    parser.add_argument('--resolution', type=int, choices=[1, 2, 3, 4, 5, 6, 7, 8, 9, 10], default=5, help='Resolution factor for the output image')

    args = parser.parse_args()

    plot_training_curves(args.data, args.Column, args.save, args.resolution)