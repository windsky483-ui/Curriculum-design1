function H = calc_entropy_gray(img)
% 计算图像灰度信息熵

img = im2double(img);

if ndims(img) == 3
    img = rgb2gray(img);
end

img = im2uint8(mat2gray(img));
counts = imhist(img, 256);
p = counts / numel(img);
p = p(p > 0);
H = -sum(p .* log2(p));
end
