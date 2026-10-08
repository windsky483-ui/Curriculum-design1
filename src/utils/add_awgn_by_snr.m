function y = add_awgn_by_snr(x, snrDb)
% 按信噪比添加高斯白噪声

x = double(x);
signalPower = mean(x(:).^2);
noisePower = signalPower / (10^(snrDb / 10));
noise = sqrt(noisePower) * randn(size(x));
y = x + noise;
end
