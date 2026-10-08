# AI驱动传输图像的无线通信系统

本项目是《通信工程综合设计 I》课程设计的实现代码，使用 MATLAB 完成图像去噪、自编码器压缩重建和 AWGN 信道传输仿真。

## 运行环境

- MATLAB R2024b 或相近版本
- Deep Learning Toolbox
- Image Processing Toolbox

## 运行顺序

1. 运行 `src/step1_dncnn_denoise.m`，测试 DnCNN 图像去噪。
2. 运行 `src/step2_train_autoencoder.m`，训练自编码器并生成 `simple_autoencoder.mat`。
3. 运行 `src/step3_jscc_awgn_simulation.m`，模拟自编码器通过 AWGN 信道传输，并计算不同信噪比下的 PSNR、MSE 和信息熵。

## 文件说明

- `src/step1_dncnn_denoise.m`：DnCNN 图像去噪。
- `src/step2_train_autoencoder.m`：训练图像压缩自编码器。
- `src/step3_jscc_awgn_simulation.m`：AWGN 信道传输与性能评估。
- `src/utils/add_awgn_by_snr.m`：按指定信噪比添加高斯白噪声。
- `src/utils/calc_entropy_gray.m`：计算图像灰度信息熵。

