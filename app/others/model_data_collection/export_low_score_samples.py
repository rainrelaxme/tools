import os
import re
import shutil
from collections import Counter

import pandas as pd


def strip_prefix(filename: str) -> str:
    """
    去掉类似于 S00027-P000- 的前缀。
    兼容不同位数的样本ID，例如 S00123、S000123 等。
    """
    if not isinstance(filename, str):
        return filename
    return re.sub(r"^S\d{5,}-P000-", "", filename)


def choose_status(status_series: pd.Series) -> str:
    """从一组状态值中选择一个最合理的（多数票，忽略空值）。"""
    values = [str(x).strip() for x in status_series if str(x).strip() not in ("", "nan", "None")]
    if not values:
        return ""
    return Counter(values).most_common(1)[0][0]


def filter_images_by_score(excel_path: str, score_column: str = "分值", name_column: str = "图片名称",
                           status_column: str = "自动标注结果", threshold: float = 0.8) -> pd.DataFrame:
    """
    读取 Excel，按图片聚合，选择分值最小的小于阈值的图片。
    若同一图片多行，则以最小分值为准，并汇总一个状态。
    返回列：图片名称（原始）、去前缀文件名、状态、最小分值。
    """
    df = pd.read_excel(excel_path, engine="openpyxl")

    if name_column not in df.columns or score_column not in df.columns:
        raise ValueError("Excel中缺少必要的列：图片名称 或 分值")

    if status_column not in df.columns:
        df[status_column] = ""

    # 分组计算每张图片的最小分值
    grouped = (
        df.groupby(name_column)
          .agg(最小分值=(score_column, "min"), 状态=(status_column, choose_status))
          .reset_index()
    )

    # 过滤分值小于阈值
    filtered = grouped[grouped["最小分值"] < threshold].copy()
    filtered["去前缀文件名"] = filtered[name_column].apply(strip_prefix)
    return filtered[[name_column, "去前缀文件名", "状态", "最小分值"]]


def ensure_dir(path: str) -> None:
    if not os.path.exists(path):
        os.makedirs(path, exist_ok=True)


def find_existing_file(root_dir: str, relative_filename: str) -> str:
    """
    在 root_dir 的 OK/NG 目录下查找文件（仅文件名匹配，不递归）
    返回找到的完整路径或空字符串。
    """
    for status in ("OK", "NG"):
        candidate = os.path.join(root_dir, status, relative_filename)
        if os.path.isfile(candidate):
            return candidate
    return ""


def copy_image_and_json(source_root: str, dest_root: str, relative_filename: str, status_hint: str = "") -> dict:
    """
    根据 relative_filename 在 source_root/OK 或 source_root/NG 下寻找并复制图片及同名 json。
    status_hint 若提供且存在对应文件，则按该状态目录复制；否则自动判断所在目录。
    返回处理结果字典。
    """
    result = {"filename": relative_filename, "status": "", "image_copied": False, "json_copied": False}

    # 先根据 hint 定位
    candidate_paths = []
    if status_hint in ("OK", "NG"):
        candidate_paths.append(os.path.join(source_root, status_hint, relative_filename))

    # 再尝试自动查找
    candidate_paths.append(find_existing_file(source_root, relative_filename))

    src_path = next((p for p in candidate_paths if p and os.path.isfile(p)), "")
    if not src_path:
        return result

    # 确认状态
    status_dir = os.path.basename(os.path.dirname(src_path))
    if status_dir not in ("OK", "NG"):
        # 容错：若不是直接在 OK/NG 下，fallback 使用 hint 或 NG
        status_dir = status_hint if status_hint in ("OK", "NG") else "NG"

    # 目标目录
    target_dir = os.path.join(dest_root, status_dir)
    ensure_dir(target_dir)

    # 复制图片
    dst_img = os.path.join(target_dir, os.path.basename(relative_filename))
    shutil.copy2(src_path, dst_img)
    result["status"] = status_dir
    result["image_copied"] = True

    # 复制同名 json（如果存在）
    json_name = os.path.splitext(os.path.basename(relative_filename))[0] + ".json"
    json_src = os.path.join(os.path.dirname(src_path), json_name)
    if os.path.isfile(json_src):
        dst_json = os.path.join(target_dir, json_name)
        shutil.copy2(json_src, dst_json)
        result["json_copied"] = True

    return result


def export_low_score_samples(excel_path: str, source_root: str, output_root: str, threshold: float = 0.8) -> pd.DataFrame:
    """
    主流程：
    1) 从 Excel 里筛选分值 < threshold 的图片（按图片聚合，处理多缺陷）。
    2) 去掉前缀后到 source_root/OK 和 source_root/NG 中查找并复制图片及 json。
    3) 输出复制结果汇总 DataFrame。
    """
    records = []
    filtered = filter_images_by_score(excel_path, threshold=threshold)

    ensure_dir(output_root)

    for _, row in filtered.iterrows():
        original_name = row["图片名称"]
        stripped_name = row["去前缀文件名"]
        status_hint = str(row["状态"]).strip()

        result = copy_image_and_json(source_root, output_root, stripped_name, status_hint=status_hint)
        records.append({
            "图片名称": original_name,
            "去前缀文件名": stripped_name,
            "状态(Excel)": status_hint,
            "分值(最小)": row["最小分值"],
            "复制到": result.get("status", ""),
            "复制图片": result.get("image_copied", False),
            "复制JSON": result.get("json_copied", False)
        })

    return pd.DataFrame(records)


def main():
    excel_path = input("请输入统计Excel路径（如 07-18_0__status.xlsx）：\n").strip()
    source_root = input("请输入源目录（含 OK/NG 子目录，如 0718OKorNG）：\n").strip()
    output_root = input("请输入输出目录（将生成 OK/NG 子目录）：\n").strip()
    threshold_str = input("请输入分值阈值（默认 0.8）：\n").strip()

    threshold = float(threshold_str) if threshold_str else 0.8

    summary = export_low_score_samples(excel_path, source_root, output_root, threshold=threshold)

    # 将结果保存到同目录
    report_path = os.path.join(output_root, "low_score_copy_result.xlsx")
    ensure_dir(output_root)
    summary.to_excel(report_path, index=False)
    print(f"完成，结果已保存：{report_path}")


if __name__ == "__main__":
    main()


