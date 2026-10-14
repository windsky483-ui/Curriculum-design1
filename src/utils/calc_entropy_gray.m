function H = calc_entropy_gray(img)
% 共同的[0,1]量化范围，避免逐图mat2gray拉伸造成比较偏差。
img = im2double(img);
if ndims(img) == 3, img = rgb2gray(img); end
counts = imhist(im2uint8(min(max(img,0),1)),256);
p = counts/numel(img);
p = p(p>0);
H = -sum(p.*log2(p));
end
