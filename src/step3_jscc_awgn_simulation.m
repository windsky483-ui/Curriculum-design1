%% 自编码器通过AWGN信道传输
clear; clc; close all;

scriptDir = fileparts(mfilename("fullpath"));
addpath(fullfile(scriptDir, "utils"));
rng(2026);

modelPath = "simple_autoencoder.mat";
if ~isfile(modelPath)
    error("Model file simple_autoencoder.mat not found. Run src/step2_train_autoencoder.m first.");
end

load(modelPath, "autoenc", "targetSize");

original = im2double(imread("cameraman.tif"));
original = imresize(original, targetSize);
if ndims(original) == 3
    original = rgb2gray(original);
end

captureNoiseSigma = 10 / 255;
captured = imnoise(original, "gaussian", 0, captureNoiseSigma^2);

try
    dncnn = denoisingNetwork("DnCNN");
    preprocessed = denoiseImage(captured, dncnn);
catch
    warning("DnCNN is unavailable. Continuing with the captured noisy image.");
    preprocessed = captured;
end

inputVector = preprocessed(:);
encoded = encode(autoenc, inputVector);

snrRange = 0:2:20;
psnrValues = zeros(size(snrRange));
mseValues = zeros(size(snrRange));
entropyValues = zeros(size(snrRange));
reconstructedImages = cell(size(snrRange));

for idx = 1:numel(snrRange)
    snrDb = snrRange(idx);
    received = add_awgn_by_snr(encoded, snrDb);
    decoded = decode(autoenc, received);
    reconstructed = reshape(decoded, targetSize);
    reconstructed = min(max(reconstructed, 0), 1);

    reconstructedImages{idx} = reconstructed;
    mseValues(idx) = immse(reconstructed, original);
    psnrValues(idx) = psnr(reconstructed, original);
    entropyValues(idx) = calc_entropy_gray(reconstructed);
end

figure("Name", "SNR performance");
tiledlayout(1, 2, "TileSpacing", "compact", "Padding", "compact");
nexttile;
plot(snrRange, psnrValues, "-o", "LineWidth", 1.5);
grid on; xlabel("SNR (dB)"); ylabel("PSNR (dB)"); title("PSNR-SNR");
nexttile;
plot(snrRange, entropyValues, "-s", "LineWidth", 1.5);
grid on; xlabel("SNR (dB)"); ylabel("Entropy (bits)"); title("Entropy-SNR");

exampleSnr = 20;
[~, exampleIndex] = min(abs(snrRange - exampleSnr));

figure("Name", "Example transmission result");
tiledlayout(1, 4, "TileSpacing", "compact", "Padding", "compact");
nexttile; imshow(original); title("Original");
nexttile; imshow(captured); title("Captured noisy");
nexttile; imshow(preprocessed); title("Preprocessed");
nexttile; imshow(reconstructedImages{exampleIndex}); title("Reconstructed");

resultTable = table(snrRange(:), mseValues(:), psnrValues(:), entropyValues(:), ...
    "VariableNames", ["SNR_dB", "MSE", "PSNR_dB", "Entropy"]);
disp(resultTable);
