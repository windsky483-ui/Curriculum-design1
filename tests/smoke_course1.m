function smoke_course1
% 合成CIFAR格式数据只检查管线，输出不代表真实CIFAR-10性能。
root = fileparts(fileparts(mfilename('fullpath')));
addpath(fullfile(root,'src'),fullfile(root,'src','utils'));
cfg = project_config();
cfg.dataDir = fullfile(root,'results','smoke','data');
cfg.outputDir = fullfile(root,'results','smoke');
cfg.modelPath = fullfile(cfg.outputDir,'simple_autoencoder.mat');
cfg.maxTrainImages = 16; cfg.maxTestImages = 8;
cfg.maxEpochs = 2; cfg.channelTrials = 2; cfg.showFigures = false;
mkdir(cfg.dataDir);
rng(41);
data = zeros(10000,3072,'uint8');
data(1:16,:) = uint8(randi([0 255],16,3072));
save(fullfile(cfg.dataDir,'data_batch_1.mat'),'data');
data(1:8,:) = uint8(randi([0 255],8,3072));
save(fullfile(cfg.dataDir,'test_batch.mat'),'data');
[~,metrics] = train_course1(cfg);
assert(height(metrics)==8);
for image = {'peppers.png','cameraman.tif'}
    cfg.testImage = image{1};
    result = simulate_course1(cfg);
    assert(height(result)==11 && all(isfinite(result.PSNR_dB_mean)));
end
fprintf('PASS: synthetic-format smoke test, RGB and grayscale, not CIFAR-10 performance.\n');
end
