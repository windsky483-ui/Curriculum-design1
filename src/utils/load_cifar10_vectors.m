function vectors = load_cifar10_vectors(dataDir, split, maxImages)
% 官方MATLAB版每行包含R/G/B平面，每平面按行优先排列。
% 转换为MATLAB img(:)列向量，保持像素位置和颜色正确。
arguments
    dataDir (1,:) char
    split (1,:) char {mustBeMember(split,{'train','test'})}
    maxImages (1,1) double {mustBeInteger,mustBePositive}
end
if strcmp(split,'train')
    names = arrayfun(@(i) sprintf('data_batch_%d.mat',i),1:5,'UniformOutput',false);
    total = 50000;
else
    names = {'test_batch.mat'}; total = 10000;
end
count = min(maxImages,total);
vectors = zeros(3072,count,'single');
offset = 0;
for i = 1:numel(names)
    path = fullfile(dataDir,names{i});
    if ~isfile(path)
        error('course1:MissingCifar10','Missing %s. See README for CIFAR-10 setup.',path);
    end
    batch = load(path,'data');
    if ~isfield(batch,'data') || ~isa(batch.data,'uint8') || ~isequal(size(batch.data),[10000 3072])
        error('course1:InvalidCifar10','Expected 10000-by-3072 uint8 data in %s.',path);
    end
    n = min(count-offset,size(batch.data,1));
    rgb = permute(reshape(batch.data(1:n,:).',32,32,3,n),[2 1 3 4]);
    vectors(:,offset+(1:n)) = reshape(single(rgb)/255,3072,n);
    offset = offset+n;
    if offset == count, break; end
end
end
