from PIL import Image
import numpy as np

# 读取图片
img = Image.open('../1.png')
img_array = np.array(img)

# 找到非白色区域（容差设为250，接近白色也算白色）
if img_array.shape[2] == 4:  # RGBA
    mask = (img_array[:, :, :3].min(axis=2) < 250) | (img_array[:, :, 3] < 255)
else:  # RGB
    mask = img_array.min(axis=2) < 250

# 获取边界
rows = np.any(mask, axis=1)
cols = np.any(mask, axis=0)

if rows.any() and cols.any():
    y1, y2 = np.where(rows)[0][[0, -1]]
    x1, x2 = np.where(cols)[0][[0, -1]]

    # 裁剪
    cropped = img.crop((x1, y1, x2+1, y2+1))

    # 放大到原始尺寸（填满画布）
    enlarged = cropped.resize((9000, 9000), Image.Resampling.LANCZOS)
    enlarged.save('../1_enlarged.png', quality=100, optimize=False)
    print(f"原始尺寸: {img.size}")
    print(f"裁剪后尺寸: {cropped.size}")
    print(f"放大后尺寸: {enlarged.size}")
    print(f"已保存为: 1_enlarged.png")
else:
    print("未检测到非白色内容")
