function y = add_awgn_by_snr(x, snrDb)
%ADD_AWGN_BY_SNR Add AWGN according to the measured signal power.
%   Y = ADD_AWGN_BY_SNR(X, SNRDB) adds zero-mean Gaussian noise so that the
%   output has approximately SNRDB dB signal-to-noise ratio.

x = double(x);
signalPower = mean(x(:).^2);
noisePower = signalPower / (10^(snrDb / 10));
noise = sqrt(noisePower) * randn(size(x));
y = x + noise;
end
