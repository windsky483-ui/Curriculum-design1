%% Step 2: Train image compression autoencoder
clear; clc; close all;

rng(2026);

digitDatasetPath = fullfile(matlabroot, "toolbox", "nnet", "nndemos", ...
    "nndatasets", "DigitDataset");
imds = imageDatastore(digitDatasetPath, ...
    "IncludeSubfolders", true, ...
    "LabelSource", "foldernames");

targetSize = [28 28];
numImages = numel(imds.Files);
trainData = zeros(prod(targetSize), numImages);

for idx = 1:numImages
    img = readimage(imds, idx);
    img = im2double(imresize(img, targetSize));
    if ndims(img) == 3
        img = rgb2gray(img);
    end
    trainData(:, idx) = img(:);
end

hiddenSize = 128;
autoenc = trainAutoencoder(trainData, hiddenSize, ...
    "MaxEpochs", 50, ...
    "L2WeightRegularization", 0.004, ...
    "SparsityRegularization", 4, ...
    "SparsityProportion", 0.15, ...
    "ScaleData", false, ...
    "EncoderTransferFunction", "satlin", ...
    "DecoderTransferFunction", "purelin");

reconstructed = predict(autoenc, trainData(:, 1:16));
save("simple_autoencoder.mat", "autoenc", "targetSize", "hiddenSize");

figure("Name", "Autoencoder reconstruction examples");
tiledlayout(4, 8, "TileSpacing", "compact", "Padding", "compact");
for idx = 1:16
    nexttile; imshow(reshape(trainData(:, idx), targetSize)); title("Input");
    nexttile; imshow(reshape(reconstructed(:, idx), targetSize)); title("Recon");
end

fprintf("Saved trained model to simple_autoencoder.mat\n");
