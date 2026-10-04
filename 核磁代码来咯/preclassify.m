
final_moral_table = randomize_blocks('E:\fMRI代码及其附属文件\核磁代码来咯\代码中用到的Excel文件\词语维度分类.1.xlsx');
    final_food_table = randomize_blocks_other('E:\fMRI代码及其附属文件\核磁代码来咯\代码中用到的Excel文件\词语维度分类.17.xlsx');
   % final_recognise_table = randomize_blocks_other('E:\fMRI代码及其附属文件\核磁代码来咯\代码中用到的Excel文件\词语维度分类.177.xlsx');
    final_food_table= renameTableColumns(final_food_table, '2');
   % final_recognise_table= renameTableColumns(final_recognise_table, '3');
    alltable = [final_moral_table,final_food_table];
    writetable(alltable, 'E:\fMRI代码及其附属文件\临时数据存放点\实验材料表.xlsx');


    function final_table = randomize_blocks(filename)
% 随机化block词语并保证每个block包含每个维度的词语
% 输入参数：
% filename - Excel文件路径
% 输出参数：
% final_table - 随机化后的表格

% 读取Excel文件
data = readtable(filename); % 读取数据
dimensions = height(data); % 维度数量
words_per_dimension = width(data); % 每个维度的单词数量

% 确保维度数量为5
if dimensions ~= 5
    error('数据必须有5个维度');
end

% 确保words_per_dimension可以被8整除
if mod(words_per_dimension, 8) ~= 0
    error('每个维度的单词数量必须能被8整除');
end

% 计算每个维度分成的行数 x
x = words_per_dimension / 8; % 每个维度分成x行8列的数据

% 假设 data 是一个 table 类型的变量（为什么要假设）
% 首先将 table 转换为数组
data_array = table2array(data);

% 初始化最终的数据表
final_table = zeros(5 * x, 8);

% 随机分配每个维度的数据
for i = 1:dimensions
    % 随机打乱数据
    dim_data = reshape(data_array(i,:), x, 8); % 使用数组进行 reshape
    % 将当前维度的数据放入最终的数据表中
    final_table((i-1)*x + 1:i*x, :) = dim_data;
end

% 将结果转换回 table 类型
final_table = array2table(final_table);

% 按行打乱数据（每一行内的数据顺序改变，但行之间的顺序不变）
final_array = table2array(final_table); % 将 table 转换为数组
for i = 1:size(final_array, 1) % 遍历每一行
    row_data = final_array(i, :); % 提取第 i 行的数据
    random_col_indices = randperm(size(final_array, 2)); % 生成随机列索引
    final_array(i, :) = row_data(random_col_indices); % 重新排列第 i 行的数据
end
final_table = array2table(final_array); % 将数组转换回 table

% 按列打乱数据（每一列内的数据顺序改变，但列之间的顺序不变）
final_array = table2array(final_table); % 将 table 转换为数组
for j = 1:size(final_array, 2) % 遍历每一列
    col_data = final_array(:, j); % 提取第 j 列的数据
    random_row_indices = randperm(size(final_array, 1)); % 生成随机行索引
    final_array(:, j) = col_data(random_row_indices); % 重新排列第 j 列的数据
end
final_table = array2table(final_array); % 将数组转换回 table
end


function final_table_other = randomize_blocks_other(filename)
% 随机化block词语并保证每个block包含每个维度的词语
% 输入参数：
% filename - Excel文件路径
% 输出参数：
% final_table_other - 随机化后的表格

% 读取Excel文件
data = readtable(filename); % 读取数据
total_words = height(data); % 总单词数量

% 确保总单词数量可以被4整除
if mod(total_words, 8) ~= 0
    error('总单词数量必须能被4整除');
end

% 计算每个block的行数 x
x = total_words / 8; % 每个block分成x行4列的数据

% 将 table 转换为数组
data_array = table2array(data);

% 随机打乱数据
data_array = data_array(randperm(total_words));

% 将随机打乱后的数据划分为x行8列
final_table_other = reshape(data_array, [8, x])'; % 转置以满足x行4列的格式

% 将结果转换回 table 类型
final_table_other = array2table(final_table_other);
end

function t = renameTableColumns(t, suffix)
    % 为 table 的每个变量名添加后缀
    t.Properties.VariableNames = strcat(t.Properties.VariableNames, '_', suffix);
end
