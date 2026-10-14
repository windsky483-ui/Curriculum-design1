function tests = test_course1_helpers
tests = functiontests(localfunctions);
end

function setupOnce(testCase)
root = fileparts(fileparts(mfilename('fullpath')));
addpath(fullfile(root,'src'),fullfile(root,'src','utils'));
testCase.TestData.tempDir = tempname;
mkdir(testCase.TestData.tempDir);
end

function teardownOnce(testCase)
% 只删除由本测试创建的MATLAB临时目录。
rmdir(testCase.TestData.tempDir,'s');
end

function testCifarPixelAndChannelOrder(testCase)
data = zeros(10000,3072,'uint8');
% CIFAR行优先像素：位置(row=2,col=3)是第35个元素。
data(1,35) = 100; data(1,1024+35) = 150; data(1,2048+35) = 200;
save(fullfile(testCase.TestData.tempDir,'data_batch_1.mat'),'data');
vectors = load_cifar10_vectors(testCase.TestData.tempDir,'train',1);
img = reshape(vectors,[32 32 3]);
verifyEqual(testCase,squeeze(img(2,3,:)),single([100;150;200])/255);
verifyEqual(testCase,squeeze(img(3,2,:)),zeros(3,1,'single'));
end

function testMissingDatasetFails(testCase)
verifyError(testCase,@() load_cifar10_vectors('missing-data','test',1),'course1:MissingCifar10');
end

function testGrayToRgb(testCase)
input = uint8(reshape(mod(0:1023,256),32,32));
img = prepare_rgb_image(input,[32 32 3]);
verifySize(testCase,img,[32 32 3]);
verifyEqual(testCase,img(:,:,1),img(:,:,2));
verifyEqual(testCase,img(:,:,1),double(input)/255);
end

function testFixedRangeEntropy(testCase)
verifyEqual(testCase,calc_entropy_gray(0.4*ones(8)),0,'AbsTol',1e-12);
verifyEqual(testCase,calc_entropy_gray([zeros(8,4),ones(8,4)]),1,'AbsTol',1e-12);
end

function testMeasuredAwgnPower(testCase)
rng(17); x = 2*ones(200000,1); y = add_awgn_by_snr(x,10);
measured = 10*log10(mean(x.^2)/mean((y-x).^2));
verifyLessThan(testCase,abs(measured-10),0.1);
verifyEqual(testCase,add_awgn_by_snr(zeros(10,1),10),zeros(10,1));
end

function testNormalizationPreservesSymbols(testCase)
x = [0.1;0.4;0.8]; scale = sqrt(mean(x.^2));
verifyEqual(testCase,mean((x/scale).^2),1,'AbsTol',1e-12);
verifyEqual(testCase,(x/scale)*scale,x,'AbsTol',1e-12);
end
