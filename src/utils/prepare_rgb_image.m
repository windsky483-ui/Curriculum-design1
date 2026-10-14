function img = prepare_rgb_image(input, targetSize)
% 灰度复制为三通道；指标参考图也使用相同的32x32分辨率。
if ischar(input) || isstring(input), input = imread(input); end
img = im2double(input);
if ismatrix(img), img = repmat(img,1,1,3); end
if ndims(img) ~= 3 || size(img,3) ~= 3
    error('course1:InvalidImage','Expected grayscale or RGB image.');
end
img = min(max(imresize(img,targetSize(1:2)),0),1);
end
