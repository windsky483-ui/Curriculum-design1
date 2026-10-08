function H = calc_entropy_gray(img)
%CALC_ENTROPY_GRAY Calculate 8-bit grayscale entropy.
%   H = CALC_ENTROPY_GRAY(IMG) converts IMG to grayscale when needed and
%   returns Shannon entropy in bits.

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
