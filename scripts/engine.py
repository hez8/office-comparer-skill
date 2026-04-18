import cv2
import numpy as np

def fast_compare_images(img1_path, img2_path, threshold):
    # 使用 OpenCV 读取，速度远快于 Pillow
    img1 = cv2.imread(img1_path)
    img2 = cv2.imread(img2_path)
    
    # 获取最大尺寸
    h, w = max(img1.shape[0], img2.shape[0]), max(img1.shape[1], img2.shape[1])
    
    # 填充大小
    a1 = np.zeros((h, w, 3), np.uint8)
    a2 = np.zeros((h, w, 3), np.uint8)
    a1[:img1.shape[0], :img1.shape[1]] = img1
    a2[:img2.shape[0], :img2.shape[1]] = img2
    
    # OpenCV 高性能逐像素差分 (使用 numpy 向量化计算)
    diff = cv2.absdiff(a1, a2)
    # 将 BGR 转为灰度以计算差异强度
    gray_diff = cv2.cvtColor(diff, cv2.COLOR_BGR2GRAY)
    
    # 差异判定
    mask_above = gray_diff > threshold
    
    # 生成可视化结果
    result = np.full((h, w, 3), 128, dtype=np.uint8) # 灰色底
    result[mask_above] = [0, 0, 255] # 差异处标记为红色 (BGR)
    
    return result, a1, a2

if __name__ == "__main__":
    import sys
    # 命令行调用接口
    res, a, b = fast_compare_images(sys.argv[1], sys.argv[2], int(sys.argv[3]))
    cv2.imwrite("diff_result.png", res)
    print("Comparison complete. Result saved to diff_result.png")
