clear;clc;
path=pwd;
PsychDefaultSetup(0);

try

    %% 被试信息
    % suborder=17;
    %MATLAB自带inputdlg函数，其作用是打开一个收集用户输入的对话框，并将输入的信息赋给一个变量
    % inputdlg中，prompt是对需要输入的信息的提示，title就是对话框标题，“[1, 50]”是输入框的维度，表示对话框的高度为一个字符，长度为50个字符。
    prompt = {'Subject Number','Gender[1 = male, 2 = female]','Age','Handeness[1 = left, 2 = right]', 'Name[initials]','Subject Order'};
    title = 'Exp infor'; % The title of the dialog boxy
    default_input = {'999','999','999','999','999','999'}; % Default input value(s)
    subinfo = inputdlg(prompt,title,[1, 50],default_input)';
    suborder = str2double(subinfo{6});
    % 读入相关材料
    tryblockorder = readcell("E:\fMRI代码及其附属文件\核磁代码来咯\代码中用到的Excel文件\被试实验次序表.xlsx");
    % 获取数列的长度
    tryblockorder = cell2mat(tryblockorder);
    wrapped_suborder = mod(suborder - 1, 30) + 1;
    tryconditionList = tryblockorder(wrapped_suborder*2-1,:);
    % 读取整个表格
    materialfullTable = readtable('E:\fMRI代码及其附属文件\临时数据存放点\实验材料表.xlsx');
    % 选择特定的列
    final_moral_table = materialfullTable(:, 1:8);
    final_food_table =materialfullTable(:, 9:16);
    moral_block_result=table2array(final_moral_table);
    moralR_try_numRows = height(moral_block_result);
    food_block_result=table2array(final_food_table);
    food_try_numRows = height(food_block_result);

    %% 实验信息配置
    try_trialnum=240;

    %%
    %% 窗口准备工作
    % 配置基本属性
    Screen('Preference', 'Verbosity', 0);
    Screen('Preference', 'SkipSyncTests', 1); %'SkipSyncTests' 偏好为 1，意味着在初始化屏幕或绘图时，Psychtoolbox 将跳过同步测试
    Screen ('Preference', 'VisualDebugLevel', 0); %'VisualDebugLevel' 偏好为 1。这个设置用于控制 Psychtoolbox 在运行时输出的视觉调试信息的级别

    %% 窗口选择
    windowSet = 1; % 0 = 1280*720; 1 =  [0, 0, 2160, 1440];
    if windowSet == 0
        % 窗口参数设置
        windowRect = [0, 0, 1280, 720]; %配置窗口大小
        [window,wsize]=Screen('OpenWindow', 0, [0,0,0], windowRect); %Screen函数第一个参数0设置调用屏幕，第二个参数[0 0 0]设置背景颜色，第三个参数设置窗口尺寸[0,0,2550,1600];调用 Screen('OpenWindow', ...) 函数时，它返回一个窗口标识符（通常是一个整数）和一个表示窗口大小的向量(写入Wsize)
        Screen('TextSize', window, 30); %配置字体大小
        frameRect_size = 3; %矩形边框宽度


        % 评分界面参数设置
        baseRect = [0 0 10 30]; % 鼠标对应刻度线位置的方块大小
        centerCorr1 = 25; % 使滑条上文本在鼠标位置正上方
        main_line_size = 5; % 主刻度线宽度
        vertical_line_size = 5; % 竖刻度线宽度
        vertical_line_height = 10; % 竖刻度线高度
        RectWidth = 4; %鼠标对应刻度线位置的方块

    elseif windowSet == 1
        % 窗口参数设置
        windowRect = [0, 0, 2160, 1440]; %配置窗口大小
        [window,wsize]=Screen('OpenWindow', 0, [0,0,0], windowRect); %Screen函数第一个参数0设置调用屏幕，第二个参数[0 0 0]设置背景颜色，第三个参数设置窗口尺寸[0,0,2550,1600];调用 Screen('OpenWindow', ...) 函数时，它返回一个窗口标识符（通常是一个整数）和一个表示窗口大小的向量(写入Wsize)
        Screen('TextSize', window, 40); %配置字体大小
        frameRect_size = 5; %矩形边框宽度


        %评分界面参数设置
        baseRect = [0 0 20 60];
        centerCorr1 = 14;
        centerCorr2 = 26;
        main_line_size = 8;
        vertical_line_size = 8;
        vertical_line_height = 16; % 竖刻度线高度
        RectWidth = 7; %鼠标对应刻度线位置的方块

    end

    [XCenter, ~] = RectCenter(wsize); %获取屏幕中心点，RectCenter用于计算一个矩形的中心点坐标
    screenWidth = wsize(3);
    screenHeight = wsize(4);

    Screen('TextFont', window, 'Microsoft YaHei'); %配置字体

    slack=Screen('GetFlipInterval',window)/2; % Screen('GetFlipInterval', window) 函数来获取指定窗口 window 的屏幕刷新间隔，以确保绘图操作能够在下一次屏幕刷新之前完成
    HideCursor; %HideCursor函数用于隐藏当前屏幕上的鼠标光标

    t=GetSecs;

    %% 数据收集
    % 练习回合数据存放
    try_picID = zeros(try_trialnum, 1);
    try_subinfo = zeros(try_trialnum,1);
    %try_pic= zeros(try_trialnum,1);
    try_block= zeros(try_trialnum,1);
    try_subADblock= zeros(try_trialnum,1);
    try_subConditionblock= zeros(try_trialnum,1);
    blocksubmergeorder = zeros(try_trialnum, 1);

    try_starttime = zeros(try_trialnum, 1);
    try_startA = zeros(try_trialnum, 1);
    try_startB = zeros(try_trialnum, 1);
    try_startC = zeros(try_trialnum, 1);
    try_startRating = zeros(try_trialnum,1);
    try_Rating = zeros(try_trialnum,1);
    try_RT_AC = zeros(try_trialnum, 1);
    try_RT_AB = zeros(try_trialnum, 1);


    try_confidencestartA = zeros(try_trialnum, 1);
    try_confidencestartB = zeros(try_trialnum, 1);
    try_confidencestartC = zeros(try_trialnum, 1);
    try_confidencestartRating = zeros(try_trialnum,1);
    try_confidenceRating = zeros(try_trialnum,1);
    try_confidenceRT_AC = zeros(try_trialnum, 1);
    try_confidenceRT_AB = zeros(try_trialnum, 1);

    for i = 1: length(subinfo)
        try_subinfo(i) = str2double(subinfo{i});
    end
    % 初始化 try_pic 为单元数组
    %try_pic = cell(try_trialnum, 1);

    showexperimentInstruction(window, "E:\fMRI代码及其附属文件\指导语\croppedinstru_2.bmp",  slack, [0, 0, 2160, 1440]);
    % 暂停 0.5 秒
    WaitSecs(0.5);
    %调整block数，1-16（16为一个block）
    try_blocktime=1;
    while try_blocktime<=16
        if tryconditionList(1,try_blocktime)==1 ||  tryconditionList(1,try_blocktime)==2 ||...
            tryconditionList(1,try_blocktime)==3 ||  tryconditionList(1,try_blocktime)==4 ||...
            tryconditionList(1,try_blocktime)==5 ||  tryconditionList(1,try_blocktime)==6 ||...
            tryconditionList(1,try_blocktime)==7 ||  tryconditionList(1,try_blocktime)==8
            %显示指导语
            showInstruction(window, "E:\fMRI代码及其附属文件\指导语\cropped_9.bmp",  slack, [0, 0, 2160, 1440]);
            %采取材料 条件编号是几，就提取第几列
            if (mod(tryconditionList(1,try_blocktime),4))==0 && tryconditionList(1,try_blocktime)<=4
                moralR_try_shuffled = moral_block_result(:,4);
            elseif (mod(tryconditionList(1,try_blocktime),4))~=0 && tryconditionList(1,try_blocktime)<=4
                moralR_try_shuffled = moral_block_result(:,(mod(tryconditionList(1,try_blocktime),4)));
            elseif (mod(tryconditionList(1,try_blocktime),4))==0 && tryconditionList(1,try_blocktime)>4
                moralR_try_shuffled = moral_block_result(:,8);
            elseif (mod(tryconditionList(1,try_blocktime),4))~=0 && tryconditionList(1,try_blocktime)>4
                moralR_try_shuffled = moral_block_result(:,((mod(tryconditionList(1,try_blocktime),4))+4));
            end
            
            %调整每个block的试次数
            for i = 1: moralR_try_numRows
                k=15*(tryconditionList(1,try_blocktime)-1)+i;
                try_image_ID = moralR_try_shuffled(i,1);
                filePath = ['E:\\fMRI代码及其附属文件\\正式试次图片\\' num2str(try_image_ID) '.bmp'];
                disp(filePath);  % 打印路径，检查是否正确
                if ~exist(filePath, 'file')
                    error('文件不存在，请检查路径和文件名: %s', filePath);
                end
                pic_image = imread(filePath);
                try_picID(k) = try_image_ID;
                pic = Screen('MakeTexture', window, pic_image);
                image_width = size(pic_image, 2);
                image_height = size(pic_image, 1);
                image_ratio = image_height / image_width;
                image_display_scale = 0.35;
                image_display_width = image_display_scale * screenWidth;
                picY = 0.18;
                Screen('DrawTexture', window, pic, [0, 0, image_width, image_height], [((1 - image_display_scale) / 2) * screenWidth, picY * screenHeight, ((1 - image_display_scale) / 2) * screenWidth + image_display_width, picY * screenHeight + image_display_width * image_ratio]);
                Screen('FrameRect', window, [255, 255, 255], [-frameRect_size + ((1 - image_display_scale) / 2) * screenWidth, -frameRect_size + picY * screenHeight, frameRect_size + ((1 - image_display_scale) / 2) * screenWidth + image_display_width, frameRect_size + picY * screenHeight + image_display_width * image_ratio], frameRect_size);
                Screen('Flip', window, t - slack);
                try_starttime(k) = GetSecs;
                WaitSecs(1.5);
                %% 道德评分部分

                lowerText = '完全没有不道德';
                upperText = '非常不道德';
                lineLength = 0.5 * screenWidth;
                halfLength = lineLength / 2;
                divider = lineLength / 6;
                lineX = XCenter;
                lineY = 0.55 * screenHeight;
                % 绘制初始图片和评分线
                Screen('DrawTexture', window, pic, [0,0,image_width,image_height],[((1-image_display_scale)/2)*screenWidth,picY*screenHeight,((1-image_display_scale)/2)*screenWidth+image_display_width,picY*screenHeight+image_display_width*image_ratio]);
                Screen('FrameRect', window, [255,255,255],[-frameRect_size+((1-image_display_scale)/2)*screenWidth, -frameRect_size+picY*screenHeight, frameRect_size+((1-image_display_scale)/2)*screenWidth+image_display_width, frameRect_size+picY*screenHeight+image_display_width*image_ratio], frameRect_size)
                Screen('DrawLine', window, [255,255,255], (lineX-halfLength-vertical_line_size/2), (lineY), (lineX+halfLength+vertical_line_size/2), (lineY), main_line_size);

                for scalepoint = -3:3
                    Screen('DrawLine', window, [255,255,255], (lineX + scalepoint*divider), (lineY), (lineX + scalepoint*divider), (lineY-vertical_line_height), vertical_line_size);
                end
                % 初始化计时与变量
                try_startA(k) = GetSecs;  % 开始呈现评分界面的时间点
                confirmed = false;
                currentRating = (randi(31)-1);
                try_startRating(k) = currentRating;
                Screen('Flip', window, t - slack);

                %% 主评分循环
                while ~confirmed
                    [keyIsDown, ~, keyCode] = KbCheck;
                    % 键盘操作
                    if keyIsDown
                        if keyCode(KbName('j')) || keyCode(KbName('J'))
                            currentRating = max(0, currentRating - 1);
                            if try_startB(k) == 0
                                try_startB(k) = GetSecs; % 第一次移动时间
                            end
                            WaitSecs(0.1);
                        elseif keyCode(KbName('k')) || keyCode(KbName('K'))
                            currentRating = min(30, currentRating + 1);
                            if try_startB(k) == 0
                                try_startB(k) = GetSecs;
                            end
                            WaitSecs(0.1);
                        elseif keyCode(KbName('f')) || keyCode(KbName('F'))
                            confirmed = true;
                            try_startC(k) = GetSecs; % 按下B的时间
                            WaitSecs(0.2);
                        end
                    end

                    % 绘制评分界面
                    ratingText = [num2str(currentRating)];
                    Screen('DrawTexture', window, pic, [0,0,image_width,image_height],[((1-image_display_scale)/2)*screenWidth,picY*screenHeight,((1-image_display_scale)/2)*screenWidth+image_display_width,picY*screenHeight+image_display_width*image_ratio]);
                    Screen('FrameRect', window, [255,255,255],[-frameRect_size+((1-image_display_scale)/2)*screenWidth, -frameRect_size+picY*screenHeight, frameRect_size+((1-image_display_scale)/2)*screenWidth+image_display_width, frameRect_size+picY*screenHeight+image_display_width*image_ratio], frameRect_size)
                    Screen('DrawLine', window, [255,255,255], (lineX-halfLength-vertical_line_size/2), (lineY), (lineX+halfLength+vertical_line_size/2), (lineY), main_line_size);

                    for scalepoint = -3:3
                        Screen('DrawLine', window, [255,255,255], (lineX + scalepoint*divider), (lineY), (lineX + scalepoint*divider), (lineY-vertical_line_height), vertical_line_size);
                    end
                    DrawFormattedText(window, double(lowerText), lineX-halfLength-0.04*screenWidth, lineY+0.12*screenHeight,255);
                    DrawFormattedText(window, double(upperText), lineX+halfLength-0.04*screenWidth, lineY+0.12*screenHeight, 255);
                    DrawFormattedText(window, double('您估计对方会在多大程度上觉得该词语所描述的行为是不道德的？'), 'center', 0.43*screenHeight, 255);
                    Screen('FillRect', window, [0,255,0], [lineX+divider*(currentRating-15)/5-RectWidth, lineY-RectWidth*3, lineX+divider*(currentRating-15)/5+RectWidth, lineY+RectWidth*3]);
                    if currentRating <=9
                        DrawFormattedText(window, ratingText, lineX-centerCorr1+divider*(currentRating-15)/5, lineY-0.04*screenHeight, [0,255,0]);
                    else
                        DrawFormattedText(window, ratingText, lineX-centerCorr2+divider*(currentRating-15)/5, lineY-0.04*screenHeight, [0,255,0]);
                    end
                    Screen('Flip', window);
                end

                %% 记录时间点与计算结果
                try_RT_AC(k) = try_startC(k) - try_startA(k); % 反应时
                if try_startB(k)~=0
                    try_RT_AB(k) = try_startB(k) - try_startA(k); % 准备移动时长
                else
                    try_RT_AB(k)=0;
                end
                try_Rating(k) = round(currentRating);

                %% 鼠标反馈屏
                if currentRating <=9
                    DrawFormattedText(window, ratingText, lineX-centerCorr1+divider*(currentRating-15)/5, lineY-0.04*screenHeight,[255,0,0]);
                else
                    DrawFormattedText(window, ratingText, lineX-centerCorr2+divider*(currentRating-15)/5, lineY-0.04*screenHeight,[255,0,0]);
                end
                DrawFormattedText(window, double('您估计对方会在多大程度上觉得该词语所描述的行为是不道德的？'),'center', 0.43*screenHeight, 255);
                Screen('DrawTexture', window, pic, [0,0,image_width,image_height],[((1-image_display_scale)/2)*screenWidth,picY*screenHeight,((1-image_display_scale)/2)*screenWidth+image_display_width,picY*screenHeight+image_display_width*image_ratio]);
                Screen('FrameRect', window, [255,255,255],[-frameRect_size+((1-image_display_scale)/2)*screenWidth, -frameRect_size+picY*screenHeight, frameRect_size+((1-image_display_scale)/2)*screenWidth+image_display_width, frameRect_size+picY*screenHeight+image_display_width*image_ratio], frameRect_size)
                Screen('DrawLine', window, [255,255,255], (lineX-halfLength-vertical_line_size/2), (lineY), (lineX+halfLength+vertical_line_size/2), (lineY), main_line_size);
                for scalepoint = -3:3
                    Screen('DrawLine', window, [255,255,255], (lineX + scalepoint*divider), (lineY), (lineX + scalepoint*divider), (lineY-vertical_line_height), vertical_line_size);
                end
                DrawFormattedText(window, double(lowerText), lineX-halfLength-0.04*screenWidth, lineY+0.12*screenHeight,255)
                DrawFormattedText(window, double(upperText), lineX+halfLength-0.04*screenWidth, lineY+0.12*screenHeight,255)
                Screen('FillRect', window, [255,0,0], [lineX+divider*(currentRating-15)/5-RectWidth,lineY-RectWidth*3,lineX+divider*(currentRating-15)/5+RectWidth,lineY+RectWidth*3]);
                Screen('Flip', window, t-slack);
                WaitSecs(0.5);

                %% 信心评分部分

                lowerText = '完全没有信心';
                upperText = '非常有信心';
                lineLength = 0.5 * screenWidth;
                halfLength = lineLength / 2;
                divider = lineLength / 6;
                lineX = XCenter;
                lineY = 0.55 * screenHeight;

                % 绘制初始图片和评分线
                Screen('DrawTexture', window, pic, [0,0,image_width,image_height],[((1-image_display_scale)/2)*screenWidth,picY*screenHeight,((1-image_display_scale)/2)*screenWidth+image_display_width,picY*screenHeight+image_display_width*image_ratio]);
                Screen('FrameRect', window, [255,255,255],[-frameRect_size+((1-image_display_scale)/2)*screenWidth, -frameRect_size+picY*screenHeight, frameRect_size+((1-image_display_scale)/2)*screenWidth+image_display_width, frameRect_size+picY*screenHeight+image_display_width*image_ratio], frameRect_size)
                Screen('DrawLine', window, [255,255,255], (lineX-halfLength-vertical_line_size/2), (lineY), (lineX+halfLength+vertical_line_size/2), (lineY), main_line_size);

                for scalepoint = -3:3
                    Screen('DrawLine', window, [255,255,255], (lineX + scalepoint*divider), (lineY), (lineX + scalepoint*divider), (lineY-vertical_line_height), vertical_line_size);
                end
                % 初始化计时与变量
                try_confidencestartA(k) = GetSecs;  % 开始呈现评分界面的时间点
                confirmed = false;
                currentRating = (randi(31)-1);
                try_confidencestartRating(k) = currentRating;
                m = 0;
                Screen('Flip', window, t - slack);

                %% 主评分循环
                while ~confirmed
                    [keyIsDown, ~, keyCode] = KbCheck;
                    if keyIsDown
                        if keyCode(KbName('j')) || keyCode(KbName('J'))
                            currentRating = max(0, currentRating - 1);
                            if try_confidencestartB(k) == 0
                                try_confidencestartB(k) = GetSecs; % 第一次移动时间
                            end
                            WaitSecs(0.1);
                        elseif keyCode(KbName('k')) || keyCode(KbName('K'))
                            currentRating = min(30, currentRating + 1);
                            if try_confidencestartB(k) == 0
                                try_confidencestartB(k) = GetSecs;
                            end
                            WaitSecs(0.1);
                        elseif keyCode(KbName('f')) || keyCode(KbName('F'))
                            confirmed = true;
                            try_confidencestartC(k) = GetSecs; % 按下B的时间
                            WaitSecs(0.2);
                        end
                    end

                    % 绘制评分界面
                    ratingText = [num2str(currentRating)];
                    Screen('DrawTexture', window, pic, [0,0,image_width,image_height],[((1-image_display_scale)/2)*screenWidth,picY*screenHeight,((1-image_display_scale)/2)*screenWidth+image_display_width,picY*screenHeight+image_display_width*image_ratio]);
                    Screen('FrameRect', window, [255,255,255],[-frameRect_size+((1-image_display_scale)/2)*screenWidth, -frameRect_size+picY*screenHeight, frameRect_size+((1-image_display_scale)/2)*screenWidth+image_display_width, frameRect_size+picY*screenHeight+image_display_width*image_ratio], frameRect_size)
                    Screen('DrawLine', window, [255,255,255], (lineX-halfLength-vertical_line_size/2), (lineY), (lineX+halfLength+vertical_line_size/2), (lineY), main_line_size);

                    for scalepoint = -3:3
                        Screen('DrawLine', window, [255,255,255], (lineX + scalepoint*divider), (lineY), (lineX + scalepoint*divider), (lineY-vertical_line_height), vertical_line_size);
                    end
                    DrawFormattedText(window, double(lowerText), lineX-halfLength-0.04*screenWidth, lineY+0.12*screenHeight, [255,165,0]);
                    DrawFormattedText(window, double(upperText), lineX+halfLength-0.04*screenWidth, lineY+0.12*screenHeight, [255,165,0]);
                    DrawFormattedText(window, double('您在多大程度上对自己初次做出的道德估计有信心？'), 'center', 0.43*screenHeight, [255,165,0]);
                    Screen('FillRect', window, [0,255,0], [lineX+divider*(currentRating-15)/5-RectWidth, lineY-RectWidth*3, lineX+divider*(currentRating-15)/5+RectWidth, lineY+RectWidth*3]);
                    if currentRating <=9
                        DrawFormattedText(window, ratingText, lineX-centerCorr1+divider*(currentRating-15)/5, lineY-0.04*screenHeight, [0,255,0]);
                    else
                        DrawFormattedText(window, ratingText, lineX-centerCorr2+divider*(currentRating-15)/5, lineY-0.04*screenHeight, [0,255,0]);
                    end
                    Screen('Flip', window);
                end

                %% 记录时间点与计算结果
                try_confidenceRT_AC(k) = try_confidencestartC(k) - try_confidencestartA(k); % 反应时
                if try_confidencestartB(k)~=0
                    try_confidenceRT_AB(k) = try_confidencestartB(k) - try_confidencestartA(k); % 准备移动时长
                else
                    try_confidenceRT_AB(k)=0;
                end
                try_confidenceRating(k) = round(currentRating);


                %% 鼠标反馈屏
                if currentRating <=9
                    DrawFormattedText(window, ratingText, lineX-centerCorr1+divider*(currentRating-15)/5, lineY-0.04*screenHeight,[255,0,0]);
                else
                    DrawFormattedText(window, ratingText, lineX-centerCorr2+divider*(currentRating-15)/5, lineY-0.04*screenHeight,[255,0,0]);
                end
                DrawFormattedText(window, double('您在多大程度上对自己初次做出的道德估计有信心？'),'center', 0.43*screenHeight, [255, 165, 0]);
                Screen('DrawTexture', window, pic, [0,0,image_width,image_height],[((1-image_display_scale)/2)*screenWidth,picY*screenHeight,((1-image_display_scale)/2)*screenWidth+image_display_width,picY*screenHeight+image_display_width*image_ratio]);
                Screen('FrameRect', window, [255,255,255],[-frameRect_size+((1-image_display_scale)/2)*screenWidth, -frameRect_size+picY*screenHeight, frameRect_size+((1-image_display_scale)/2)*screenWidth+image_display_width, frameRect_size+picY*screenHeight+image_display_width*image_ratio], frameRect_size)
                Screen('DrawLine', window, [255,255,255], (lineX-halfLength-vertical_line_size/2), (lineY), (lineX+halfLength+vertical_line_size/2), (lineY), main_line_size);
                for scalepoint = -3:3
                    Screen('DrawLine', window, [255,255,255], (lineX + scalepoint*divider), (lineY), (lineX + scalepoint*divider), (lineY-vertical_line_height), vertical_line_size);
                end
                DrawFormattedText(window, double(lowerText), lineX-halfLength-0.04*screenWidth, lineY+0.12*screenHeight, [255, 165, 0])
                DrawFormattedText(window, double(upperText), lineX+halfLength-0.04*screenWidth, lineY+0.12*screenHeight, [255, 165, 0])
                Screen('FillRect', window, [255,0,0], [lineX+divider*(currentRating-15)/5-RectWidth,lineY-RectWidth*3,lineX+divider*(currentRating-15)/5+RectWidth,lineY+RectWidth*3]);
                Screen('Flip', window, t-slack);
                WaitSecs(0.5);
                try_block(k)=tryconditionList(1,try_blocktime);
                try_subConditionblock(k)= 1;
                if mod(tryconditionList(1,try_blocktime),4)==0
                    try_subADblock(k)=4;
                else
                    try_subADblock(k)=mod(tryconditionList(1,try_blocktime),4);
                end
                blocksubmergeorder(k)=try_blocktime;
            end
        elseif tryconditionList(1,try_blocktime)==9 ||  tryconditionList(1,try_blocktime)==10 ||tryconditionList(1,try_blocktime)==11 ||  tryconditionList(1,try_blocktime)==12 || tryconditionList(1,try_blocktime)==13 ||  tryconditionList(1,try_blocktime)==14 ||tryconditionList(1,try_blocktime)==15 ||  tryconditionList(1,try_blocktime)==16
            %指导语
            showInstruction(window, "E:\fMRI代码及其附属文件\指导语\cropped_10.bmp",  slack, [0, 0, 2160, 1440]);
            %材料
            if (mod(tryconditionList(1,try_blocktime),4))==0 && tryconditionList(1,try_blocktime)<=12
                food_try_shuffled = food_block_result(:,4);
            elseif (mod(tryconditionList(1,try_blocktime),4))~=0 && tryconditionList(1,try_blocktime)<=12
                food_try_shuffled = food_block_result(:,(mod(tryconditionList(1,try_blocktime),4)));
            elseif (mod(tryconditionList(1,try_blocktime),4))==0 && tryconditionList(1,try_blocktime)>12
                food_try_shuffled = food_block_result(:,8);
            elseif (mod(tryconditionList(1,try_blocktime),4))~=0 && tryconditionList(1,try_blocktime)>12
                food_try_shuffled = food_block_result(:,((mod(tryconditionList(1,try_blocktime),4))+4));
            end

            %食物试次
            %1-15
            for i = 1:food_try_numRows
                k=15*(tryconditionList(1,try_blocktime)-1)+i;
                try_image_ID = food_try_shuffled(i,1);
                filePath = ['E:\\fMRI代码及其附属文件\\正式试次图片\\' num2str(try_image_ID) '.bmp'];
                disp(filePath);  % 打印路径，检查是否正确
                if ~exist(filePath, 'file')
                    error('文件不存在，请检查路径和文件名: %s', filePath);
                end
                pic_image = imread(filePath);
                try_picID(k) = try_image_ID;
                pic = Screen('MakeTexture', window, pic_image);
                image_width = size(pic_image, 2);
                image_height = size(pic_image, 1);
                image_ratio = image_height / image_width;
                image_display_scale = 0.35;
                image_display_width = image_display_scale * screenWidth;
                picY = 0.18;
                Screen('DrawTexture', window, pic, [0, 0, image_width, image_height], [((1 - image_display_scale) / 2) * screenWidth, picY * screenHeight, ((1 - image_display_scale) / 2) * screenWidth + image_display_width, picY * screenHeight + image_display_width * image_ratio]);
                Screen('FrameRect', window, [255, 255, 255], [-frameRect_size + ((1 - image_display_scale) / 2) * screenWidth, -frameRect_size + picY * screenHeight, frameRect_size + ((1 - image_display_scale) / 2) * screenWidth + image_display_width, frameRect_size + picY * screenHeight + image_display_width * image_ratio], frameRect_size);
                Screen('Flip', window, t - slack);
                try_starttime(k) = GetSecs;
                WaitSecs(1.5);
                %% 食物评分部分

                lowerText = '完全没有不喜欢';
                upperText = '非常不喜欢';
                lineLength = 0.5 * screenWidth;
                halfLength = lineLength / 2;
                divider = lineLength / 6;
                lineX = XCenter;
                lineY = 0.55 * screenHeight;

                % 绘制初始图片和评分线
                Screen('DrawTexture', window, pic, [0,0,image_width,image_height],[((1-image_display_scale)/2)*screenWidth,picY*screenHeight,((1-image_display_scale)/2)*screenWidth+image_display_width,picY*screenHeight+image_display_width*image_ratio]);
                Screen('FrameRect', window, [255,255,255],[-frameRect_size+((1-image_display_scale)/2)*screenWidth, -frameRect_size+picY*screenHeight, frameRect_size+((1-image_display_scale)/2)*screenWidth+image_display_width, frameRect_size+picY*screenHeight+image_display_width*image_ratio], frameRect_size)
                Screen('DrawLine', window, [255,255,255], (lineX-halfLength-vertical_line_size/2), (lineY), (lineX+halfLength+vertical_line_size/2), (lineY), main_line_size);

                for scalepoint = -3:3
                    Screen('DrawLine', window, [255,255,255], (lineX + scalepoint*divider), (lineY), (lineX + scalepoint*divider), (lineY-vertical_line_height), vertical_line_size);
                end
                % 初始化计时与变量
                try_startA(k) = GetSecs;  % 开始呈现评分界面的时间点
                confirmed = false;
                currentRating = (randi(31)-1);
                try_startRating(k) = currentRating;
                Screen('Flip', window, t - slack);

                %% 主评分循环
                while ~confirmed
                    [keyIsDown, ~, keyCode] = KbCheck;
                    % 键盘操作
                    if keyIsDown
                        if keyCode(KbName('j')) || keyCode(KbName('J'))
                            currentRating = max(0, currentRating - 1);
                            if try_startB(k) == 0
                                try_startB(k) = GetSecs; % 第一次移动时间
                            end
                            WaitSecs(0.1);
                        elseif keyCode(KbName('k')) || keyCode(KbName('K'))
                            currentRating = min(30, currentRating + 1);
                            if try_startB(k) == 0
                                try_startB(k) = GetSecs;
                            end
                            WaitSecs(0.1);
                        elseif keyCode(KbName('f')) || keyCode(KbName('F'))
                            confirmed = true;
                            try_startC(k) = GetSecs; % 按下B的时间
                            WaitSecs(0.2);
                        end
                    end

                    % 绘制评分界面
                    ratingText = [num2str(currentRating)];
                    Screen('DrawTexture', window, pic, [0,0,image_width,image_height],[((1-image_display_scale)/2)*screenWidth,picY*screenHeight,((1-image_display_scale)/2)*screenWidth+image_display_width,picY*screenHeight+image_display_width*image_ratio]);
                    Screen('FrameRect', window, [255,255,255],[-frameRect_size+((1-image_display_scale)/2)*screenWidth, -frameRect_size+picY*screenHeight, frameRect_size+((1-image_display_scale)/2)*screenWidth+image_display_width, frameRect_size+picY*screenHeight+image_display_width*image_ratio], frameRect_size)
                    Screen('DrawLine', window, [255,255,255], (lineX-halfLength-vertical_line_size/2), (lineY), (lineX+halfLength+vertical_line_size/2), (lineY), main_line_size);

                    for scalepoint = -3:3
                        Screen('DrawLine', window, [255,255,255], (lineX + scalepoint*divider), (lineY), (lineX + scalepoint*divider), (lineY-vertical_line_height), vertical_line_size);
                    end
                    DrawFormattedText(window, double(lowerText), lineX-halfLength-0.04*screenWidth, lineY+0.12*screenHeight,255);
                    DrawFormattedText(window, double(upperText), lineX+halfLength-0.04*screenWidth, lineY+0.12*screenHeight, 255);
                    DrawFormattedText(window, double('您估计对方会在多大程度上不喜欢该词语所描述的食物？'), 'center', 0.43*screenHeight, 255);
                    Screen('FillRect', window, [0,255,0], [lineX+divider*(currentRating-15)/5-RectWidth, lineY-RectWidth*3, lineX+divider*(currentRating-15)/5+RectWidth, lineY+RectWidth*3]);
                    if currentRating <=9
                        DrawFormattedText(window, ratingText, lineX-centerCorr1+divider*(currentRating-15)/5, lineY-0.04*screenHeight, [0,255,0]);
                    else
                        DrawFormattedText(window, ratingText, lineX-centerCorr2+divider*(currentRating-15)/5, lineY-0.04*screenHeight, [0,255,0]);
                    end
                    Screen('Flip', window);
                end

                %% 记录时间点与计算结果
                try_RT_AC(k) = try_startC(k) - try_startA(k); % 反应时
                if try_startB(k)~=0
                    try_RT_AB(k) = try_startB(k) - try_startA(k); % 准备移动时长
                else
                    try_RT_AB(k)=0;
                end
                try_Rating(k) = round(currentRating);

                %% 鼠标反馈屏
                if currentRating <=9
                    DrawFormattedText(window, ratingText, lineX-centerCorr1+divider*(currentRating-15)/5, lineY-0.04*screenHeight,[255,0,0]);
                else
                    DrawFormattedText(window, ratingText, lineX-centerCorr2+divider*(currentRating-15)/5, lineY-0.04*screenHeight,[255,0,0]);
                end
                DrawFormattedText(window, double('您估计对方会在多大程度上不喜欢该词语所描述的食物？'),'center', 0.43*screenHeight, 255);
                Screen('DrawTexture', window, pic, [0,0,image_width,image_height],[((1-image_display_scale)/2)*screenWidth,picY*screenHeight,((1-image_display_scale)/2)*screenWidth+image_display_width,picY*screenHeight+image_display_width*image_ratio]);
                Screen('FrameRect', window, [255,255,255],[-frameRect_size+((1-image_display_scale)/2)*screenWidth, -frameRect_size+picY*screenHeight, frameRect_size+((1-image_display_scale)/2)*screenWidth+image_display_width, frameRect_size+picY*screenHeight+image_display_width*image_ratio], frameRect_size)
                Screen('DrawLine', window, [255,255,255], (lineX-halfLength-vertical_line_size/2), (lineY), (lineX+halfLength+vertical_line_size/2), (lineY), main_line_size);
                for scalepoint = -3:3
                    Screen('DrawLine', window, [255,255,255], (lineX + scalepoint*divider), (lineY), (lineX + scalepoint*divider), (lineY-vertical_line_height), vertical_line_size);
                end
                DrawFormattedText(window, double(lowerText), lineX-halfLength-0.04*screenWidth, lineY+0.12*screenHeight,255)
                DrawFormattedText(window, double(upperText), lineX+halfLength-0.04*screenWidth, lineY+0.12*screenHeight,255)
                Screen('FillRect', window, [255,0,0], [lineX+divider*(currentRating-15)/5-RectWidth,lineY-RectWidth*3,lineX+divider*(currentRating-15)/5+RectWidth,lineY+RectWidth*3]);
                Screen('Flip', window, t-slack);
                WaitSecs(0.5);

                %% 信心评分部分

                lowerText = '完全没有信心';
                upperText = '非常有信心';
                lineLength = 0.5 * screenWidth;
                halfLength = lineLength / 2;
                divider = lineLength / 6;
                lineX = XCenter;
                lineY = 0.55 * screenHeight;

                % 绘制初始图片和评分线
                Screen('DrawTexture', window, pic, [0,0,image_width,image_height],[((1-image_display_scale)/2)*screenWidth,picY*screenHeight,((1-image_display_scale)/2)*screenWidth+image_display_width,picY*screenHeight+image_display_width*image_ratio]);
                Screen('FrameRect', window, [255,255,255],[-frameRect_size+((1-image_display_scale)/2)*screenWidth, -frameRect_size+picY*screenHeight, frameRect_size+((1-image_display_scale)/2)*screenWidth+image_display_width, frameRect_size+picY*screenHeight+image_display_width*image_ratio], frameRect_size)
                Screen('DrawLine', window, [255,255,255], (lineX-halfLength-vertical_line_size/2), (lineY), (lineX+halfLength+vertical_line_size/2), (lineY), main_line_size);

                for scalepoint = -3:3
                    Screen('DrawLine', window, [255,255,255], (lineX + scalepoint*divider), (lineY), (lineX + scalepoint*divider), (lineY-vertical_line_height), vertical_line_size);
                end
                % 初始化计时与变量
                try_confidencestartA(k) = GetSecs;  % 开始呈现评分界面的时间点
                confirmed = false;
                currentRating = (randi(31)-1);
                try_confidencestartRating(k) = currentRating;
                m = 0;
                Screen('Flip', window, t - slack);

                %% 主评分循环
                while ~confirmed
                    [keyIsDown, ~, keyCode] = KbCheck;
                    if keyIsDown
                        if keyCode(KbName('j')) || keyCode(KbName('J'))
                            currentRating = max(0, currentRating - 1);
                            if try_confidencestartB(k) == 0
                                try_confidencestartB(k) = GetSecs; % 第一次移动时间
                            end
                            WaitSecs(0.1);
                        elseif keyCode(KbName('k')) || keyCode(KbName('K'))
                            currentRating = min(30, currentRating + 1);
                            if try_confidencestartB(k) == 0
                                try_confidencestartB(k) = GetSecs;
                            end
                            WaitSecs(0.1);
                        elseif keyCode(KbName('f')) || keyCode(KbName('F'))
                            confirmed = true;
                            try_confidencestartC(k) = GetSecs; % 按下B的时间
                            WaitSecs(0.2);
                        end
                    end

                    % 绘制评分界面
                    ratingText = [num2str(currentRating)];
                    Screen('DrawTexture', window, pic, [0,0,image_width,image_height],[((1-image_display_scale)/2)*screenWidth,picY*screenHeight,((1-image_display_scale)/2)*screenWidth+image_display_width,picY*screenHeight+image_display_width*image_ratio]);
                    Screen('FrameRect', window, [255,255,255],[-frameRect_size+((1-image_display_scale)/2)*screenWidth, -frameRect_size+picY*screenHeight, frameRect_size+((1-image_display_scale)/2)*screenWidth+image_display_width, frameRect_size+picY*screenHeight+image_display_width*image_ratio], frameRect_size)
                    Screen('DrawLine', window, [255,255,255], (lineX-halfLength-vertical_line_size/2), (lineY), (lineX+halfLength+vertical_line_size/2), (lineY), main_line_size);

                    for scalepoint = -3:3
                        Screen('DrawLine', window, [255,255,255], (lineX + scalepoint*divider), (lineY), (lineX + scalepoint*divider), (lineY-vertical_line_height), vertical_line_size);
                    end
                    DrawFormattedText(window, double(lowerText), lineX-halfLength-0.04*screenWidth, lineY+0.12*screenHeight, [255,165,0]);
                    DrawFormattedText(window, double(upperText), lineX+halfLength-0.04*screenWidth, lineY+0.12*screenHeight, [255,165,0]);
                    DrawFormattedText(window, double('您在多大程度上对自己初次做出的食物估计有信心？'), 'center', 0.43*screenHeight, [255,165,0]);
                    Screen('FillRect', window, [0,255,0], [lineX+divider*(currentRating-15)/5-RectWidth, lineY-RectWidth*3, lineX+divider*(currentRating-15)/5+RectWidth, lineY+RectWidth*3]);
                    if currentRating <=9
                        DrawFormattedText(window, ratingText, lineX-centerCorr1+divider*(currentRating-15)/5, lineY-0.04*screenHeight, [0,255,0]);
                    else
                        DrawFormattedText(window, ratingText, lineX-centerCorr2+divider*(currentRating-15)/5, lineY-0.04*screenHeight, [0,255,0]);
                    end
                    Screen('Flip', window);
                end

                %% 记录时间点与计算结果
                try_confidenceRT_AC(k) = try_confidencestartC(k) - try_confidencestartA(k); % 反应时
                if try_confidencestartB(k)~=0
                    try_confidenceRT_AB(k) = try_confidencestartB(k) - try_confidencestartA(k); % 准备移动时长
                else
                    try_confidenceRT_AB(k)=0;
                end
                try_confidenceRating(k) = round(currentRating);


                %% 鼠标反馈
                if currentRating <=9
                    DrawFormattedText(window, ratingText, lineX-centerCorr1+divider*(currentRating-15)/5, lineY-0.04*screenHeight,[255,0,0]);
                else
                    DrawFormattedText(window, ratingText, lineX-centerCorr2+divider*(currentRating-15)/5, lineY-0.04*screenHeight,[255,0,0]);
                end
                DrawFormattedText(window, double('您在多大程度上对自己初次做出的食物估计有信心？'),'center', 0.43*screenHeight, [255, 165, 0]);
                Screen('DrawTexture', window, pic, [0,0,image_width,image_height],[((1-image_display_scale)/2)*screenWidth,picY*screenHeight,((1-image_display_scale)/2)*screenWidth+image_display_width,picY*screenHeight+image_display_width*image_ratio]);
                Screen('FrameRect', window, [255,255,255],[-frameRect_size+((1-image_display_scale)/2)*screenWidth, -frameRect_size+picY*screenHeight, frameRect_size+((1-image_display_scale)/2)*screenWidth+image_display_width, frameRect_size+picY*screenHeight+image_display_width*image_ratio], frameRect_size)
                Screen('DrawLine', window, [255,255,255], (lineX-halfLength-vertical_line_size/2), (lineY), (lineX+halfLength+vertical_line_size/2), (lineY), main_line_size);
                for scalepoint = -3:3
                    Screen('DrawLine', window, [255,255,255], (lineX + scalepoint*divider), (lineY), (lineX + scalepoint*divider), (lineY-vertical_line_height), vertical_line_size);
                end
                DrawFormattedText(window, double(lowerText), lineX-halfLength-0.04*screenWidth, lineY+0.12*screenHeight, [255, 165, 0])
                DrawFormattedText(window, double(upperText), lineX+halfLength-0.04*screenWidth, lineY+0.12*screenHeight, [255, 165, 0])
                Screen('FillRect', window, [255,0,0], [lineX+divider*(currentRating-15)/5-RectWidth,lineY-RectWidth*3,lineX+divider*(currentRating-15)/5+RectWidth,lineY+RectWidth*3]);
                Screen('Flip', window, t-slack);
                WaitSecs(0.5);
                try_block(k)=tryconditionList(1,try_blocktime);
                try_subConditionblock(k)= 2;
                if mod(tryconditionList(1,try_blocktime),4)==0
                    try_subADblock(k)=4;
                else
                    try_subADblock(k)=mod(tryconditionList(1,try_blocktime),4);
                end
                blocksubmergeorder(k)=try_blocktime;
            end
        end
        if mod(try_blocktime,2)==0 && try_blocktime<16
            breakconfirmed = false;
            breaktimeoutShown = false;
            tiptime = GetSecs;

            while ~breakconfirmed
                [keyIsDown, ~, keyCode] = KbCheck;

                % 检查是否按下Q键
                if keyIsDown && (keyCode(KbName('q')) || keyCode(KbName('Q')))
                    % 显示继续任务提示
                    DrawFormattedText(window, double('将继续任务，请做好准备'), 'center', 'center', 255);
                    Screen('Flip', window);
                    WaitSecs(2);
                    breakconfirmed = true;
                    break;  % 跳出当前循环
                end

                % 检查是否超时（1分钟）
                if ~breaktimeoutShown && (GetSecs - tiptime) > 60
                    % 显示超时提示并等待按键
                    while ~breakconfirmed
                        DrawFormattedText(window, double('休息时间结束，请按"Q"键继续任务'), 'center', 'center', [255,0,0]);
                        Screen('Flip', window);

                        [keyIsDown, ~, keyCode] = KbCheck;
                        if keyIsDown && (keyCode(KbName('q')) || keyCode(KbName('Q')))
                            DrawFormattedText(window, double('将继续任务，请做好准备'), 'center', 'center', 255);
                            Screen('Flip', window);
                            WaitSecs(2);
                            breakconfirmed = true;
                            break;
                        end
                    end
                    breaktimeoutShown = true;
                end

                % 显示休息提示（仅在未超时且未确认时）
                if ~breakconfirmed && ~breaktimeoutShown
                    DrawFormattedText(window, double('您可以休息1分钟（如需跳过休息，可按"Q"键继续任务）'), 'center', 'center', 255);
                    Screen('Flip', window);
                end
            end
        end
        try_blocktime=try_blocktime+1;
    end

    %% 结束界面
    ending_Key = KbName('y');
    while 1
        DrawFormattedText(window, double('您已完成初次估计的全部任务，请联系实验人员'),'center', 'center', [255,255,255]);
        Screen('Flip',window,t-slack); %t-slack在下一个屏幕刷新周期的开始时更新窗口内容
        [~,~,keyCode]=KbCheck;
        if keyCode(ending_Key)
            break;
        end
    end


    Screen('CloseAll');
    Screen('Close');
    fclose('all');
    %% 数据存放
    filename = "E:\fMRI代码及其附属文件\临时数据存放点\FTtryrating.xlsx";
    writematrix(try_Rating, filename, 'WriteMode', 'overwrite');

    % 获取 subInfo 的长度
    try_subinfo_length = length(try_subinfo);

    % 如果 subInfo 的长度小于 try_trialnum，则复制最后一行
    if try_subinfo_length < try_trialnum
        % 计算需要复制的次数
        num_repeats = try_trialnum - try_subinfo_length;

        % 复制最后一行
        try_subinfo = [try_subinfo; repmat(try_subinfo(end, :), num_repeats, 1)];
    end


    if ~all([length(try_picID),  length(try_starttime),length(try_startA), length(try_startB),length(try_startC),length(try_startRating),length(try_Rating), length(try_RT_AC), length(try_RT_AB), length(try_confidencestartA), length(try_confidencestartB),length(try_confidencestartC),length(try_confidencestartRating), length(try_confidenceRating), length(try_confidenceRT_AC), length(try_confidenceRT_AB),length(try_subADblock),length(try_subConditionblock),length(try_block)] == try_trialnum)
        error('数组长度不一致，请检查数据收集部分的代码');
    end
    rating_try_data = [try_subinfo, try_picID, try_starttime, try_startA,try_startB,try_startC,try_startRating, try_Rating,try_RT_AC,try_RT_AB,try_confidencestartA,try_confidencestartB,try_confidencestartC,try_confidencestartRating,try_confidenceRating,try_confidenceRT_AC,try_confidenceRT_AB,blocksubmergeorder,try_subADblock,try_subConditionblock,try_block];
    variableNames = {'subInfo', 'picID','showwordtime','ScreenStartTime', 'MouseMoveTime', 'ClickTime', 'randomtry_startRating','Rating','reacttime','premovetime','ConfiScreenStartTime', 'ConfiMouseMoveTime', 'ConfiClickTime', 'randomtry_confidencestartRating','confidenceRating','confi_reacttime','confi_premovetime','blockorder','try_subADblock','try_subConditionblock','try_block'}; % 列标题
    rating_try_result = table(try_subinfo, try_picID, try_starttime, try_startA,try_startB,try_startC,try_startRating, try_Rating,try_RT_AC,try_RT_AB,try_confidencestartA,try_confidencestartB,try_confidencestartC,try_confidencestartRating,try_confidenceRating,try_confidenceRT_AC,try_confidenceRT_AB,blocksubmergeorder,try_subADblock,try_subConditionblock,try_block,'VariableNames', variableNames);

    subnum = subinfo{1};  % 从 cell 里取编号（可能是数字或字符串）
    baseDir = "E:\fMRI代码及其附属文件\实验数据\first_trydata\";

    % ⚙️ 转成字符串并拼接
    filename = fullfile(baseDir, sprintf('%s.xlsx', num2str(subnum)));

    % 检查路径是否正确
    disp(['文件将被保存到：', filename]);

    % 保存文件
    try
        writetable(rating_try_result, filename);
        disp(['文件已成功保存到：', filename]);
    catch ME
        error('保存文件时出错：%s', ME.message);
    end



catch exception
    fprintf('主程序存在错误，请检查！\n');
    fprintf('警告: 转义字符 ''\'' 无效。有关支持的特殊字符，请参阅 ''doc sprintf''。\n');
    fprintf('错误位置：%s (第 %d 行)\n', exception.stack(1).file, exception.stack(1).line);
    fprintf('错误信息：%s\n', exception.message);
    Screen('CloseAll');
    Screen('Close');
    fclose('all');
end



function showexperimentInstruction(window, filePath, slack, windowRect)
disp(filePath);  % 打印路径，检查是否正确
if ~exist(filePath, 'file')
    error('文件不存在，请检查路径和文件名: %s', filePath);
end
introductionStart_image = imread(filePath);
introductionStart = Screen('MakeTexture', window, introductionStart_image);
introduction_width = size(introductionStart_image, 2);
introduction_height = size(introductionStart_image, 1);

% 初始化按键状态
keyIsDown = false;

while ~keyIsDown
    Screen('DrawTexture', window, introductionStart, [0, 0, introduction_width, introduction_height], windowRect);
    Screen('Flip', window);  % 刷新屏幕
    [~, ~, keyCode] = KbCheck;
    if keyCode(KbName('q')) || keyCode(KbName('Q'))
        disp('Q 键被按下，进入下一个指导语界面');
        keyIsDown = true;  % 设置按键标志为 true
    end
    WaitSecs(slack);  % 控制刷新频率
end
end

function showInstruction(window, filePath, slack, windowRect)
disp(filePath);  % 打印路径，检查是否正确
if ~exist(filePath, 'file')
    error('文件不存在，请检查路径和文件名: %s', filePath);
end
introductionStart_image = imread(filePath);
introductionStart = Screen('MakeTexture', window, introductionStart_image);
introduction_width = size(introductionStart_image, 2);
introduction_height = size(introductionStart_image, 1);

% 初始化按键状态
keyIsDown = false;

while ~keyIsDown
    Screen('DrawTexture', window, introductionStart, [0, 0, introduction_width, introduction_height], windowRect);
    Screen('Flip', window);  % 刷新屏幕
    [~, ~, ~] = KbCheck;
    keyIsDown = true;  % 设置按键标志为 true
    % 暂停4秒
    WaitSecs(1);
    WaitSecs(slack);  % 控制刷新频率
end
end
