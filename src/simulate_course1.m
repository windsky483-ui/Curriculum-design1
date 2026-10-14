function resultTable = simulate_course1(cfg)
% 自编码器+AWGN仿真；无噪声训练不等同于噪声感知联合训练。
addpath(fullfile(fileparts(mfilename('fullpath')),'utils'));
rng(cfg.seed);
validateattributes(cfg.channelTrials,{'numeric'},{'scalar','integer','positive'});
model = load(cfg.modelPath,'autoenc','targetSize','hiddenSize','trainingInfo');
if ~isequal(model.targetSize,cfg.targetSize) || model.hiddenSize ~= cfg.hiddenSize
    error('course1:ModelMismatch','Retrain with the current RGB/hidden-size configuration.');
end
dncnn = denoisingNetwork('DnCNN');
original = prepare_rgb_image(cfg.testImage,cfg.targetSize);
captured = imnoise(original,'gaussian',0,cfg.captureNoiseSigma^2);
preprocessed = denoise_rgb_dncnn(captured,dncnn);
encoded = double(encode(model.autoenc,preprocessed(:)));
symbolRms = sqrt(mean(encoded(:).^2));
if symbolRms == 0
    error('course1:ZeroSymbolPower','Encoder output has zero power; check model training.');
end
transmitted = encoded/symbolRms; % 单位功率；接收端恢复编码尺度
cleanReconstruction = min(max(reshape(decode(model.autoenc,encoded),cfg.targetSize),0),1);
nSnr = numel(cfg.snrRange);
mseTrials = zeros(nSnr,cfg.channelTrials);
psnrTrials = zeros(nSnr,cfg.channelTrials);
entropyTrials = zeros(nSnr,cfg.channelTrials);
examples = cell(nSnr,1);
for i = 1:nSnr
    for trial = 1:cfg.channelTrials
        received = add_awgn_by_snr(transmitted,cfg.snrRange(i))*symbolRms;
        reconstruction = min(max(reshape(decode(model.autoenc,received),cfg.targetSize),0),1);
        mseTrials(i,trial) = immse(reconstruction,original);
        psnrTrials(i,trial) = psnr(reconstruction,original);
        entropyTrials(i,trial) = calc_entropy_gray(reconstruction);
        if trial == 1, examples{i} = reconstruction; end
    end
end
resultTable = table(cfg.snrRange(:),mean(mseTrials,2),mean(psnrTrials,2), ...
    std(psnrTrials,0,2),mean(entropyTrials,2), ...
    'VariableNames',{'SNR_dB','MSE_mean','PSNR_dB_mean','PSNR_dB_std','Entropy_mean'});
disp(resultTable);
fprintf('Clean-link PSNR: %.4f dB; original entropy: %.4f bits\n', ...
    psnr(cleanReconstruction,original),calc_entropy_gray(original));
if ~isfolder(cfg.outputDir), mkdir(cfg.outputDir); end
writetable(resultTable,fullfile(cfg.outputDir,'snr_metrics.csv'));
save(fullfile(cfg.outputDir,'simulation.mat'),'cfg','resultTable','mseTrials', ...
    'psnrTrials','entropyTrials','original','captured','preprocessed', ...
    'cleanReconstruction','examples','symbolRms');
if cfg.showFigures
    performanceFigure = figure('Name','SNR performance');
    tiledlayout(1,2,'TileSpacing','compact','Padding','compact');
    nexttile; errorbar(cfg.snrRange,mean(psnrTrials,2),std(psnrTrials,0,2),'-o');
    grid on; xlabel('SNR (dB)'); ylabel('PSNR (dB)'); title('Mean +/- standard deviation');
    nexttile; plot(cfg.snrRange,mean(entropyTrials,2),'-s');
    grid on; xlabel('SNR (dB)'); ylabel('Grayscale entropy (bits)');
    exportgraphics(performanceFigure,fullfile(cfg.outputDir,'snr_performance.png'));
    [~,exampleIndex] = min(abs(cfg.snrRange-20));
    exampleFigure = figure('Name','Example transmission result');
    tiledlayout(1,5,'TileSpacing','compact','Padding','compact');
    nexttile; imshow(original); title('Reference');
    nexttile; imshow(captured); title('Captured noisy');
    nexttile; imshow(preprocessed); title('DnCNN');
    nexttile; imshow(cleanReconstruction); title('No channel noise');
    nexttile; imshow(examples{exampleIndex}); title(sprintf('%g dB, trial 1',cfg.snrRange(exampleIndex)));
    exportgraphics(exampleFigure,fullfile(cfg.outputDir,'reconstruction.png'));
end
end
