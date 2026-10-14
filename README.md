# AI驱动传输图像的无线通信系统

《通信工程综合设计 I》个人课程设计，使用 MATLAB R2024b 完成本地图像采集噪声模拟、预训练 DnCNN 去噪、CIFAR-10 自编码器压缩重建、AWGN 信道仿真及 PSNR/MSE/灰度信息熵评估。

| 项目 | 当前配置 |
|---|---|
| 训练数据 | 官方 CIFAR-10 MATLAB 格式，50,000张训练图像 |
| 独立测试数据 | 官方测试批次，10,000张，不参与训练 |
| 图像表示 | 32×32×3 RGB，MATLAB `img(:)` 展平为3072维 |
| 编码维数 | 256，可设置128/512进行对比 |
| 训练 | `trainAutoencoder`，50 epochs，`satlin` / `purelin`，`ScaleData=false` |
| 单独去噪实验 | `cameraman.tif`，采集噪声标准差25/255 |
| 完整链路 | 默认 `peppers.png`，采集噪声标准差10/255 |
| 信道 | 实值 AWGN，SNR=0:2:20 dB，每点20次独立信道重复 |
| 评估 | 同分辨率参考图上的MSE、PSNR，固定量化范围的灰度熵 |

### 实现边界

本实现是**自编码器图像压缩重建 + 潜在表示AWGN传输**的课程仿真，可用于讨论JSCC基本思想。`trainAutoencoder`在干净图像上训练，AWGN在仿真时加入，不参与反向传播，因此不应把它描述为已经完成噪声感知的端到端联合训练，也没有实现独立的调制、纠错编码或比特级BER测量。

报告中的约13.5dB曲线、约19dB示例和约20.73dB测试均值属于不同记录，不能从本次代码修改推定可复现。程序记录实际训练配置并计算指标，不内置或保证这些数值。灰度熵只描述灰度分布，不单独证明细节恢复。

## 环境和数据准备

- MATLAB R2024b
- Deep Learning Toolbox、Image Processing Toolbox
- `denoisingNetwork('DnCNN')`所需的预训练网络支持包（若提示缺失，请按MATLAB提示安装）

CIFAR-10不是 MATLAB 安装时保证自带的数据。下载[官方 CIFAR-10 MATLAB 数据包](https://www.cs.toronto.edu/~kriz/cifar-10-matlab.tar.gz)，解压到仓库的 `data/`，使目录如下：

```text
data/cifar-10-batches-mat/data_batch_1.mat
data/cifar-10-batches-mat/data_batch_2.mat
data/cifar-10-batches-mat/data_batch_3.mat
data/cifar-10-batches-mat/data_batch_4.mat
data/cifar-10-batches-mat/data_batch_5.mat
data/cifar-10-batches-mat/test_batch.mat
```

也可在仓库根目录的MATLAB命令窗口执行：

```matlab
mkdir('data');
archive = websave(fullfile('data','cifar-10-matlab.tar.gz'), ...
    'https://www.cs.toronto.edu/~kriz/cifar-10-matlab.tar.gz');
untar(archive,'data');
```

网络不通时手动下载后解压。加载器检查MAT文件形状和类型，并处理CIFAR行优先与MATLAB列优先的排列差异，避免训练图像转置或颜色错位。

## 运行

按顺序执行：

1. `src/step1_dncnn_denoise.m`：独立DnCNN去噪验证。
2. `src/step2_train_autoencoder.m`：CIFAR-10训练、独立测试及本地图像无信道噪声重建。
3. `src/step3_jscc_awgn_simulation.m`：采集噪声、逐通道DnCNN、编码、AWGN、解码、指标和图像。

训练配置在 `src/project_config.m`；模型保存到 `models/simple_autoencoder.mat`，指标和图像保存到 `results/`。脚本和模型路径不依赖当前工作目录，旧28×28/128维模型不会被当作RGB模型使用。网络缺失时明确报错，不跳过去噪模块。

输入灰度图会复制成RGB三通道以匹配训练模型；这是兼容输入，不是单独训练灰度网络。RGB逐通道使用预训练灰度DnCNN，不是另行训练的彩色DnCNN。

编码符号先按RMS归一化为单位平均功率，加入AWGN后再恢复尺度；接收端使用发送端尺度作为仿真已知量，没有计入传送尺度的开销。SNR定义使用符号平均总功率（包含直流分量）。每个SNR输出平均PSNR、PSNR标准差、平均MSE和平均灰度熵，示例图为第1次信道重复，并不保证每个点严格单调。

完整训练可能耗时且占用较多内存。`trainAutoencoder`采用该接口自身的优化过程，不提供 `MiniBatchSize` 参数。加载器使用single以减少数据存储，测试预测按256张分块；需要调试时可减小 `maxTrainImages` 和 `maxEpochs`，但必须将输出明确标为调试结果，不能称为50,000张/50轮完整实验。调试时会选取训练集前N张，不能用于无偏的模型比较。

## 排查重建模糊

先比较原图、去噪图、无信道噪声重建和有信道噪声重建：

- 无信道噪声也模糊：检查32×32降采样、训练/测试分布、训练轮次、模型容量及压缩维数。MSE目标与单隐藏层容量也可能使输出更平滑。
- 仅加入信道后恶化：检查符号功率与噪声公式，并考虑当前模型没有在含噪潜在表示上训练。
- 对比128/256/512时使用同一数据划分、随机种子和测试条件，分别保存模型、配置及实际指标。

## 验证

在仓库根目录运行：

```matlab
r = runtests('tests/test_course1_helpers.m');
assert(all([r.Passed]));
addpath('tests');
smoke_course1;
```

辅助测试覆盖CIFAR像素/通道排列、灰度输入、固定范围信息熵、AWGN实测功率和归一化尺度。`smoke_course1`以合成CIFAR格式数据做16张/2轮训练、8张独立预测，并跑RGB和灰度AWGN链路；只验证代码路径，不代表真实CIFAR-10训练性能。数据、模型和运行结果不提交到Git仓库。
