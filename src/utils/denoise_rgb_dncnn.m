function output = denoise_rgb_dncnn(input, net)
% 预训练DnCNN接受单通道输入，因此RGB逐通道处理。
% 网络缺失时报告错误，不隐式跳过去噪。
output = zeros(size(input),'like',input);
for c = 1:size(input,3)
    output(:,:,c) = denoiseImage(input(:,:,c),net);
end
output = min(max(output,0),1);
end
