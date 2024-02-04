%% 1.初始化
clear
close all
clc
format short  % 精确到小数点后4位，format long是精确到小数点后15位

%% 2.加载模型

% 加载模型
load('Model_F10.mat', 'net', 'inputps', 'outputps');

%% 3.读取数据

%读取原始数据和标题
[~, ~, raw] = xlsread('0_input_xnyy_3F_Final.xlsx');
original_data = raw(2:end,:); % 原始数据（排除标题行）
titles = raw(1,:); % 原始数据的标题行

% 提取输入数据并进行归一化
input_A = cell2mat(original_data(:,1:6))'; % 提取输入数据并转置
inputn = mapstd('apply', input_A, inputps); % 输入数据归一化

%% 4.进行预测
an = sim(net, inputn); % 使用训练好的网络进行预测

%% 5.反归一化预测结果
predicted_output = mapstd('reverse', an, outputps); % 反归一化以获得实际的输出值
predicted_output_transposed = predicted_output'; %转置

%% 6.准备写入Excel的数据
% 将原始输入数据与预测的输出数据合并
final_data = [original_data, num2cell(predicted_output_transposed)];

%% 7.添加标题
new_titles = [titles, {'超过步行率', '超过时间率', '迷茫次数'}];
final_data_with_titles = [new_titles; final_data];

%% 计算寻路影响指数
Y1 = predicted_output_transposed(:, 1); % 提取预测结果的第一个变量（Y1）
Y2 = predicted_output_transposed(:, 2); % 提取预测结果的第二个变量（Y2）
Y3 = predicted_output_transposed(:, 3); % 提取预测结果的第三个变量（Y3）
wayfinding_impact = Y1 * 0.27 + Y2 * 0.13 + Y3 * 0.60; % 计算寻路影响指数
floor_wayfinding_impact = mean(wayfinding_impact); % 计算所有样本的寻路影响指数的算数平均值

%% 添加数据至excel
final_data_with_titles(:, end + 1) = {'路径寻路影响指数'}; % 添加列标题
final_data_with_titles(2:end, end) = num2cell(wayfinding_impact); % 添加数据

final_data_with_titles(1, end + 1) = {'层平均寻路影响指数'}; % 添加列标题
final_data_with_titles{2, end} = floor_wayfinding_impact  % 添加层平均寻路影响指数的值到第二行
%% 写入Excel文件

filename = 'Assess_xnyy_3F_Final.xlsx';
xlswrite(filename, final_data_with_titles);

disp('预测完成，数据已写入文件');
