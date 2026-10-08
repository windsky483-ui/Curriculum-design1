%% DnCNN图像去噪
clear; clc; close all;

addpath(fullfile(fileparts(mfilename("fullpath")), "utils"));
rng(2026);

net = denoisingNetwork("DnCNN");

original = im2double(imread("cameraman.tif"));
noiseSigma = 25 / 255;
noisy = imnoise(original, "gaussian", 0, noiseSigma^2);
denoised = denoiseImage(noisy, net);

psnrNoisy = psnr(noisy, original);
psnrDenoised = psnr(denoised, original);

entropyOriginal = calc_entropy_gray(original);
entropyNoisy = calc_entropy_gray(noisy);
entropyDenoised = calc_entropy_gray(denoised);

fprintf("Noisy PSNR: %.2f dB\n", psnrNoisy);
fprintf("Denoised PSNR: %.2f dB\n", psnrDenoised);
fprintf("Entropy original/noisy/denoised: %.4f / %.4f / %.4f\n", ...
    entropyOriginal, entropyNoisy, entropyDenoised);

figure("Name", "DnCNN image denoising");
tiledlayout(1, 3, "TileSpacing", "compact", "Padding", "compact");
nexttile; imshow(original); title("Original");
nexttile; imshow(noisy); title("Noisy");
nexttile; imshow(denoised); title("DnCNN denoised");
