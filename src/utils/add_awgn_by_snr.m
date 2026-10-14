function y = add_awgn_by_snr(x, snrDb)
% 实值符号平均总功率(含直流分量)，不是去均值方差。
validateattributes(x,{'numeric'},{'real','finite','nonempty'});
validateattributes(snrDb,{'numeric'},{'real','finite','scalar'});
x = double(x);
noisePower = mean(x(:).^2)/10^(snrDb/10);
y = x+sqrt(noisePower)*randn(size(x));
end
