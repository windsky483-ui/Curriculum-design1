function [autoenc, testMetrics] = train_course1(cfg)
% 训练集拟合，官方测试集独立评估，不参与参数优化。
addpath(fullfile(fileparts(mfilename('fullpath')),'utils'));
rng(cfg.seed);
assert(isequal(cfg.targetSize,[32 32 3]),'CIFAR-10 requires [32 32 3].');
trainData = load_cifar10_vectors(cfg.dataDir,'train',cfg.maxTrainImages);
fprintf('CIFAR-10: %d training images, 3072 -> %d -> 3072, %d epochs\n', ...
    size(trainData,2),cfg.hiddenSize,cfg.maxEpochs);
autoenc = trainAutoencoder(trainData,cfg.hiddenSize, ...
    'MaxEpochs',cfg.maxEpochs,'L2WeightRegularization',0.004, ...
    'SparsityRegularization',4,'SparsityProportion',0.15, ...
    'ScaleData',false,'EncoderTransferFunction','satlin', ...
    'DecoderTransferFunction','purelin','ShowProgressWindow',false);
trainingInfo = struct('dataset','CIFAR-10','numTrainImages',size(trainData,2), ...
    'epochs',cfg.maxEpochs,'seed',cfg.seed,'noiseAwareTraining',false);
targetSize = cfg.targetSize; hiddenSize = cfg.hiddenSize;
folder = fileparts(cfg.modelPath);
if ~isfolder(folder), mkdir(folder); end
save(cfg.modelPath,'autoenc','targetSize','hiddenSize','trainingInfo','-v7.3');
clear trainData;

testData = load_cifar10_vectors(cfg.dataDir,'test',cfg.maxTestImages);
count = size(testData,2); mseValues = zeros(count,1);
% 分块预测减少临时内存；trainAutoencoder不提供小批量训练参数。
for first = 1:256:count
    last = min(first+255,count);
    reference = double(testData(:,first:last));
    reconstructed = min(max(double(predict(autoenc,testData(:,first:last))),0),1);
    mseValues(first:last) = mean((reconstructed-reference).^2,1).';
end
psnrValues = 10*log10(1./mseValues);
testMetrics = table((1:count).',mseValues,psnrValues, ...
    'VariableNames',{'ImageIndex','MSE','PSNR_dB'});
if ~isfolder(cfg.outputDir), mkdir(cfg.outputDir); end
writetable(testMetrics,fullfile(cfg.outputDir,'cifar10_test_metrics.csv'));
fprintf('Independent CIFAR-10 test: %d images, mean PSNR %.4f dB\n',count,mean(psnrValues));
fprintf('Saved model: %s\n',cfg.modelPath);
reference = prepare_rgb_image(cfg.testImage,cfg.targetSize);
reconstruction = reshape(predict(autoenc,reference(:)),cfg.targetSize);
reconstruction = min(max(reconstruction,0),1);
fprintf('Clean test image PSNR: %.4f dB\n',psnr(reconstruction,reference));
if cfg.showFigures
    figure('Name','Independent image reconstruction'); tiledlayout(1,2);
    nexttile; imshow(reference); title('Reference 32x32 RGB');
    nexttile; imshow(reconstruction); title('Clean autoencoder reconstruction');
end
end
