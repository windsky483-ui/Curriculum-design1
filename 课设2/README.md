# 基于AI的通信信号调制方式识别系统

## 项目简介

本项目实现了基于人工智能的通信信号调制方式自动识别系统，支持 **9种数字调制信号** (2ASK, 4ASK, 2FSK, 4FSK, BPSK, QPSK, 8PSK, 16QAM, 64QAM) 的自动识别。

识别准确率：**93.21%**（超过90%的设计要求）

## 快速开始

### 环境要求
- Python 3.8+
- 依赖库：numpy, scipy, matplotlib, scikit-learn

### 安装
```bash
pip install -r requirements.txt
```

### 运行
```bash
# 完整流程：训练模型 + 启动GUI
python main.py

# 仅训练模型
python main.py --train

# 仅启动GUI（需已有模型文件）
python main.py --gui

# 对比所有模型性能
python main.py --compare
```

## 项目结构

```
ModulationRecognition/
├── main.py                  # 主入口
├── config.py                # 系统参数配置
├── signal_generator.py      # 信号生成模块 (ASK/FSK/PSK/QAM)
├── feature_extraction.py    # 特征提取模块 (32维特征)
├── dataset_builder.py       # 数据集构建与预处理
├── ai_model.py              # SVM模型 (RBF核 + 概率校准)
├── database.py              # SQLite数据库模块
├── gui_app.py               # Tkinter可视化GUI
├── requirements.txt         # 依赖列表
├── data/                    # 数据与模型文件
│   ├── best_model.pkl       # 训练好的最佳模型
│   └── modulation_recognition.db  # 数据库
└── report/                  # 课程设计报告
    └── 课程设计报告.md
```

## 支持的调制类型

| 类型 | 描述 | 调制方式 |
|------|------|----------|
| 2ASK | 二进制幅移键控 | 幅度调制 |
| 4ASK | 四进制幅移键控 | 幅度调制 |
| 2FSK | 二进制频移键控 | 频率调制 |
| 4FSK | 四进制频移键控 | 频率调制 |
| BPSK | 二进制相移键控 | 相位调制 |
| QPSK | 四进制相移键控 | 相位调制 |
| 8PSK | 八进制相移键控 | 相位调制 |
| 16QAM | 16正交幅度调制 | 幅度+相位 |
| 64QAM | 64正交幅度调制 | 幅度+相位 |

## 特征体系 (32维)

- **瞬时特征 (14维)**：γ_max, σ_ap, σ_dp, σ_aa, σ_af, P, 幅度峰度/偏度, 频率统计等
- **频谱特征 (8维)**：频谱质心, 带宽, 峰度, 偏度, 对称性, 平坦度, 峰均比
- **高阶累积量 (10维)**：C20, C21, C40, C41, C42, C60, C61, C62, C63, C80

## 模型性能

| 模型 | 准确率 | F1分数 | 训练耗时 |
|------|--------|--------|----------|
| SVM (RBF核) | 93.21% | 93.20% | 0.162s |

## 创新点

1. **多SNR混合数据增强**：在5~30dB宽范围内训练，提升鲁棒性
2. **高阶累积量特征融合**：利用对高斯噪声免疫的统计特性进行识别
3. **SVM算法深度优化**：RBF核函数 + GridSearchCV超参数搜索 + 概率校准
4. **端到端工程实现**：信号生成→特征提取→SVM训练→GUI可视化

## 技术栈

- **信号处理**：NumPy, SciPy
- **机器学习**：Scikit-learn
- **可视化**：Matplotlib + Tkinter
- **数据库**：SQLite
