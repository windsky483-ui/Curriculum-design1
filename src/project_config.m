function cfg = project_config()
% 按报告4.1与简历采用RGB实验，路径不依赖当前工作目录。
root = fileparts(fileparts(mfilename('fullpath')));
cfg.root = root;
cfg.dataDir = fullfile(root,'data','cifar-10-batches-mat');
cfg.modelPath = fullfile(root,'models','simple_autoencoder.mat');
cfg.outputDir = fullfile(root,'results');
cfg.targetSize = [32 32 3];
cfg.hiddenSize = 256; % 可改为128或512比较压缩维数
cfg.maxEpochs = 50;
cfg.maxTrainImages = 50000; % 减小仅用于调试，不代表完整实验
cfg.maxTestImages = 10000;
cfg.captureNoiseSigma = 10/255;
cfg.snrRange = 0:2:20;
cfg.channelTrials = 20; % 每个SNR重复独立信道噪声，固定采集图像
cfg.testImage = 'peppers.png'; % 可换成本地灰度或彩色图
cfg.seed = 2026;
cfg.showFigures = true;
end
