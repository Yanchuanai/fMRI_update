% 
% jitter_data=generate_hybrid_jitter(30,1);
% function jitter_data = generate_hybrid_jitter(trial_num, sub_num)
%     rng(sub_num);
%     jitter_min = 1.0;
%     jitter_max = 3.0;
% 
%     % 生成网格点
%     grid_points = linspace(jitter_min, jitter_max, trial_num + 2);
%     grid_points = grid_points(2:end-1);  % 去掉两端点
% 
%     % 添加小扰动
%     perturbation = (jitter_max - jitter_min) / (trial_num * 2) * (rand(trial_num, 1) - 0.5);
%     jitter_data = grid_points(:) + perturbation;
% 
%     % 确保在边界内
%     jitter_data = min(max(jitter_data, jitter_min), jitter_max);
% 
%     % 随机打乱
%     jitter_data = jitter_data(randperm(trial_num));
% end
% 
% % 2. 复制8遍
% replicated_jitter = repmat(jitter_data, 8, 1);
% 
% % 3. 每30行内部打乱
% numBlocks = 8;
% shuffled_jitter = zeros(240, 1);
% 
% for block = 1:numBlocks
%     startIdx = (block-1)*30 + 1;
%     endIdx = block*30;
%     current_block = replicated_jitter(startIdx:endIdx);
% 
%     % 在块内打乱
%     randomOrder = randperm(30);
%     shuffled_jitter(startIdx:endIdx) = current_block(randomOrder);
% end
% 
% % 4. 保存结果
% filename ="D:\想法的实施\MATLAB_dm\核磁代码来咯\try_AD_jitter.xlsx";
% writematrix(shuffled_jitter, filename);
% 
% % 5. 验证
% fprintf('原始30个值平均值: %.3f\n', mean(jitter_data));
% fprintf('复制8遍后平均值: %.3f\n', mean(replicated_jitter));
% fprintf('分块打乱后平均值: %.3f\n', mean(shuffled_jitter));
% fprintf('每块平均值:\n');
% for block = 1:numBlocks
%     block_mean = mean(shuffled_jitter((block-1)*30+1:block*30));
%     fprintf('  块%d: %.3f\n', block, block_mean);
% end


jitter_data=generate_hybrid_jitter(8,1);
function jitter_data = generate_hybrid_jitter(trial_num, sub_num)
    rng(sub_num);
    jitter_min =2.0;
    jitter_max = 5.0;
    
    % 生成网格点
    grid_points = linspace(jitter_min, jitter_max, trial_num + 2);
    grid_points = grid_points(2:end-1);  % 去掉两端点
    
    % 添加小扰动
    perturbation = (jitter_max - jitter_min) / (trial_num * 2) * (rand(trial_num, 1) - 0.5);
    jitter_data = grid_points(:) + perturbation;
    
    % 确保在边界内
    jitter_data = min(max(jitter_data, jitter_min), jitter_max);
    
    % 随机打乱
    jitter_data = jitter_data(randperm(trial_num));
end
% 
% % 2. 复制8遍
% replicated_jitter = repmat(jitter_data, 8, 1);
% 
% % 3. 每30行内部打乱
% numBlocks = 8;
% shuffled_jitter = zeros(240, 1);
% 
% for block = 1:numBlocks
%     startIdx = (block-1)*30 + 1;
%     endIdx = block*30;
%     current_block = replicated_jitter(startIdx:endIdx);
% 
%     % 在块内打乱
%     randomOrder = randperm(30);
%     shuffled_jitter(startIdx:endIdx) = current_block(randomOrder);
% end

% 4. 保存结果
filename ="D:\想法的实施\MATLAB_dm\核磁代码来咯\train_random_numbers.xlsx";
writematrix(jitter_data, filename);

% 5. 验证
fprintf('原始30个值平均值: %.3f\n', mean(jitter_data));
fprintf('复制8遍后平均值: %.3f\n', mean(replicated_jitter));
fprintf('分块打乱后平均值: %.3f\n', mean(shuffled_jitter));
fprintf('每块平均值:\n');
for block = 1:numBlocks
    block_mean = mean(shuffled_jitter((block-1)*30+1:block*30));
    fprintf('  块%d: %.3f\n', block, block_mean);
end