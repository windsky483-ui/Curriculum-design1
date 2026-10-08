# AI驱动传输图像的无线通信系统

本仓库对应《通信工程综合设计 I》课程设计，主题为“AI驱动的智能通信系统”。项目使用 MATLAB 实现图像去噪、自编码器联合信源信道编码、AWGN 信道传输仿真，并通过 PSNR、MSE 和信息熵评价重建效果。

## 项目内容

- `src/step1_dncnn_denoise.m`：使用预训练 DnCNN 对图像加噪和去噪，并计算 PSNR、信息熵。
- `src/step2_train_autoencoder.m`：训练图像压缩自编码器，并保存模型。
- `src/step3_jscc_awgn_simulation.m`：加载模型，模拟完整图像通信链路，绘制 SNR 性能曲线。
- `src/utils/calc_entropy_gray.m`：灰度信息熵计算函数。
- `src/utils/add_awgn_by_snr.m`：按指定 SNR 添加 AWGN 噪声函数。
- `docs/report_outline.md`：课程设计报告结构摘要。

## 运行环境

- MATLAB R2024b 或相近版本
- Deep Learning Toolbox
- Image Processing Toolbox

首次运行建议按以下顺序执行：

```matlab
run("src/step1_dncnn_denoise.m")
run("src/step2_train_autoencoder.m")
run("src/step3_jscc_awgn_simulation.m")
```

`step2_train_autoencoder.m` 会生成 `simple_autoencoder.mat`。该文件属于训练产物，已在 `.gitignore` 中排除，重新运行脚本即可生成。

## 设计流程

1. 对输入图像添加采集噪声，模拟图像采集过程中的高斯噪声干扰。
2. 使用 DnCNN 网络进行发送端预处理，降低采集噪声对后续编码的影响。
3. 将图像归一化并展平，输入自编码器编码器得到低维符号向量。
4. 在符号向量上传输 AWGN 信道噪声，模拟不同 SNR 条件下的无线信道。
5. 通过解码器恢复图像，计算 PSNR、MSE 和灰度信息熵。
6. 绘制 SNR 与重建质量之间的关系曲线。

## 指标说明

- PSNR：用于衡量重建图像与原始图像之间的峰值信噪比。
- MSE：用于衡量像素级均方误差。
- 信息熵：基于 256 级灰度直方图，反映图像灰度分布的离散程度。

## 说明

本项目中的代码框架依据课程设计报告整理，AI 工具用于辅助方案设计、代码搭建和调试解释，最终结果需要以 MATLAB 实际运行结果为准。
