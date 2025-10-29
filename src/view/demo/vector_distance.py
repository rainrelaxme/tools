import os

import numpy as np
from onnxruntime import (SessionOptions, ExecutionMode, GraphOptimizationLevel)
from optimum.onnxruntime import ORTModelForFeatureExtraction
from transformers import AutoTokenizer

model = r"D:\Code\Config\bge-large-zh-v1.5"
# 初始化ONNX Runtime选项
ort_session_options = SessionOptions()
ort_session_options.intra_op_num_threads = 4  # 等于CPU物理核心数
ort_session_options.execution_mode = ExecutionMode.ORT_PARALLEL  # 启用并行
ort_session_options.graph_optimization_level = GraphOptimizationLevel.ORT_ENABLE_ALL
# 全局模型和分词器
tokenizer = AutoTokenizer.from_pretrained(model)  # 初始化一个预训练的分词器
onnx_model = ORTModelForFeatureExtraction.from_pretrained(
    model,
    session_options=ort_session_options
)
os.environ["TOKENIZERS_PARALLELISM"] = "false"  # 禁用Hugging Face Tokenizers 库的并行处理行为


def encode_onnx(texts):
    """
    将文本数据转换为向量表示（文本嵌入/embedding）
    """
    inputs = tokenizer(texts, padding=True, truncation=True, max_length=512, return_tensors="np")
    # 移除token_type_ids（如果存在）
    if 'token_type_ids' in inputs:
        del inputs['token_type_ids']
    outputs = onnx_model(**inputs)
    embeddings = outputs.last_hidden_state[:, 0, :]  # 取[CLS] token
    embeddings = embeddings / np.linalg.norm(embeddings, axis=1, keepdims=True)  # 归一化
    return embeddings.astype(np.float32)


if __name__ == "__main__":
    # 假设有两个向量
    # query
    query = "碰伤"
    data = "类别：结构。外观分类：。问题描述：碰伤"

    vec_a = encode_onnx([query])
    print("******vec_a*******", vec_a)
    vec_b = encode_onnx([data])
    print("******vec_b*******", vec_b)

    # 计算 L2 距离（欧氏距离）
    l2_distance = np.linalg.norm(vec_a - vec_b)
    print(f"L2 Distance: {l2_distance}")

    # 计算 IP（内积）
    ip_distance = np.dot(vec_a, vec_b)
    print(f"IP Distance: {ip_distance}")

    # # 计算 COSINE 相似度（需要先对向量进行 L2 归一化）
    # def cosine_similarity(a, b):
    #     a_norm = a / np.linalg.norm(a)
    #     b_norm = b / np.linalg.norm(b)
    #     return np.dot(a_norm, b_norm)

    # cosine_sim = cosine_similarity(vec_a, vec_b)
    # print(f"Cosine Similarity: {cosine_sim}")
