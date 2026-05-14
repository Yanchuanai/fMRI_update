clear;clc;
path=pwd;
PsychDefaultSetup(0);

try
    %% 被试信息（被试顺序，需要改）
    suborder=17;
    %MATLAB自带inputdlg函数，其作用是打开一个收集用户输入的对话框，并将输入的信息赋给一个变量
    % inputdlg中，prompt是对需要输入的信息的提示，title就是对话框标题，“[1, 50]”是输入框的维度，表示对话框的高度为一个字符，长度为50个字符。
    prompt = {'Subject Number','Gender[1 = male, 2 = female]','Age','Handeness[1 = left, 2 = right]', 'Name[initials]'};
    title = 'Exp infor'; % The title of the dialog boxy
    default_input = {'999','999','999','999','999'}; % Default input value(s)
    subinfo = inputdlg(prompt,title,[1, 50],default_input)';

    %% 读入相关材料
    jitter_try_shuffled = readmatrix ("E:\fMRI代码及其附属文件\核磁代码来咯\代码中用到的Excel文件\try_random_numbers.xlsx");
    try_AD_jitter=readmatrix ("E:\fMRI代码及其附属文件\核磁代码来咯\代码中用到的Excel文件\try_AD_jitter.xlsx");
    FTtryrating=readmatrix ("E:\fMRI代码及其附属文件\临时数据存放点\FTtryrating.xlsx");
    tryblockorder = readcell("E:\fMRI代码及其附属文件\核磁代码来咯\代码中用到的Excel文件\被试实验次序表.xlsx");
    % 获取数列的长度
    tryblockorder = cell2mat(tryblockorder);
    tryconditionList = tryblockorder(suborder*2-1,:);
    plusminusrandom = tryblockorder(suborder*2,:);
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
        centerCorr = 25; % 使滑条上文本在鼠标位置正上方
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
    fixCrossDimPix = 20; % 十字臂的一半长度（总长度为40像素）
    lineWidthPix = 4;    % 线条宽度（像素）

    [xCenter, yCenter] = RectCenter(Screen('Rect', window));

    % 横纵坐标点（相对中心）
    xCoords = [-fixCrossDimPix fixCrossDimPix 0 0];
    yCoords = [0 0 -fixCrossDimPix fixCrossDimPix];
    allCoords = [xCoords; yCoords];

    %按键设置
    makesureKey = KbName('2@');
    leftKey = KbName('3#');
    rightKey = KbName('4$');


    %% 数据收集
    % 练习回合数据存放
    try_picID = zeros(try_trialnum, 1);
    try_subinfo = zeros(try_trialnum,1);
    try_block= zeros(try_trialnum,1);
    try_subADblock= zeros(try_trialnum,1);
    try_subConditionblock= zeros(try_trialnum,1);
    try_AD_Rating= zeros(try_trialnum,1);
    blocksubmergeorder= zeros(try_trialnum,1);
    AD_suffered= zeros(try_trialnum,1);
    ADjittershow = zeros(try_trialnum, 1);
    trailjittershow = zeros(try_trialnum, 1);
    cueshow = zeros(try_trialnum, 1);
    feedbackshow = zeros(try_trialnum, 1);
    readyshow = zeros(try_trialnum, 1);
    relaxshow = zeros(try_trialnum, 1);
    overtimeshow = zeros(try_trialnum, 1);
    relaxshowtime=0;


    try_starttime = zeros(try_trialnum, 1);
    try_startA = zeros(try_trialnum, 1);
    try_startB = zeros(try_trialnum, 1);
    try_startC = zeros(try_trialnum, 1);
    try_startRating = zeros(try_trialnum,1);
    try_Rating = zeros(try_trialnum,1);
    try_isTimeout = zeros(try_trialnum, 1);
    try_RT_AC = zeros(try_trialnum, 1);
    try_RT_AB = zeros(try_trialnum, 1);
    try_feedbackTime = zeros(try_trialnum, 1);

    for i = 1: length(subinfo)
        try_subinfo(i) = str2double(subinfo{i});
    end
    % 获取建议人参考图片的原始尺寸
    imagePath1 = ['E:\\fMRI代码及其附属文件\\反转小人图片\\yellow' num2str(4) '.png'];
    originalImage1 = imread(imagePath1);
    imageHeight1 = size(originalImage1, 1);
    imageWidth1 = size(originalImage1, 2);

    wait_for_trigger_advanced
    readyshowtime=GetSecs;
    showInstruction(window, "E:\fMRI代码及其附属文件\指导语\croppedinstru_5.bmp",  slack, [0, 0, 2160, 1440]);
    %block数 =16时为一个block，=16时，block为1
    try_blocktime=1;
    while try_blocktime<=16
        if plusminusrandom(1,try_blocktime)==1%等于1意味着这个block里面8负7正；等于2意味着7负8正
            onesAndTwos = [ones(1, 8), 2 * ones(1, 7)];
        else
            onesAndTwos = [ones(1, 7), 2 * ones(1, 8)];
        end
        blockmp=onesAndTwos(randperm(15));
        if tryconditionList(1,try_blocktime)==1 ||  tryconditionList(1,try_blocktime)==2 ||tryconditionList(1,try_blocktime)==3 ||  tryconditionList(1,try_blocktime)==4 || tryconditionList(1,try_blocktime)==5 ||  tryconditionList(1,try_blocktime)==6 ||tryconditionList(1,try_blocktime)==7 ||  tryconditionList(1,try_blocktime)==8
            cueshowtime = GetSecs;
            %显示指导语
            if tryconditionList(1,try_blocktime)==1 || tryconditionList(1,try_blocktime)==5
                showInstruction(window, "E:\fMRI代码及其附属文件\指导语\cropped_1.bmp",  slack, [0, 0, 2160, 1440]);
            elseif tryconditionList(1,try_blocktime)==2 || tryconditionList(1,try_blocktime)==6
                showInstruction(window, "E:\fMRI代码及其附属文件\指导语\cropped_2.bmp",  slack, [0, 0, 2160, 1440]);
            elseif tryconditionList(1,try_blocktime)==3 || tryconditionList(1,try_blocktime)==7
                showInstruction(window, "E:\fMRI代码及其附属文件\指导语\cropped_3.bmp",  slack, [0, 0, 2160, 1440]);
            elseif tryconditionList(1,try_blocktime)==4 || tryconditionList(1,try_blocktime)==8
                showInstruction(window, "E:\fMRI代码及其附属文件\指导语\cropped_4.bmp",  slack, [0, 0, 2160, 1440]);
            end
            %采取材料
            if (mod(tryconditionList(1,try_blocktime),4))==0 && tryconditionList(1,try_blocktime)<=4
                moralR_try_shuffled = moral_block_result(:,4);
            elseif (mod(tryconditionList(1,try_blocktime),4))~=0 && tryconditionList(1,try_blocktime)<=4
                moralR_try_shuffled = moral_block_result(:,(mod(tryconditionList(1,try_blocktime),4)));
            elseif (mod(tryconditionList(1,try_blocktime),4))==0 && tryconditionList(1,try_blocktime)>4
                moralR_try_shuffled = moral_block_result(:,8);
            elseif (mod(tryconditionList(1,try_blocktime),4))~=0 && tryconditionList(1,try_blocktime)>4
                moralR_try_shuffled = moral_block_result(:,((mod(tryconditionList(1,try_blocktime),4))+4));
            end

            for i = 1: moralR_try_numRows
                k=15*(tryconditionList(1,try_blocktime)-1)+i;
                j=15*(try_blocktime-1)+i;
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
                lowerText = '完全没有不道德';
                upperText = '非常不道德';
                lineLength = 0.5 * screenWidth;
                halfLength = lineLength / 2;
                divider = lineLength / 6;
                lineX = XCenter;
                lineY = 0.55 * screenHeight;
                y_mouse = lineY;

                currentRating=FTtryrating(k);
                try_RatingText = [num2str(FTtryrating(k))];
                if FTtryrating(k) <=9
                    DrawFormattedText(window, try_RatingText, lineX-centerCorr1+divider*(FTtryrating(k)-15)/5, lineY-0.04*screenHeight, [255,0,0]);
                else
                    DrawFormattedText(window, try_RatingText, lineX-centerCorr2+divider*(FTtryrating(k)-15)/5, lineY-0.04*screenHeight, [255,0,0]);
                end
                DrawFormattedText(window, double('您的道德初次估计值为'),'center', 0.43*screenHeight, 255);
                Screen('DrawTexture', window, pic, [0,0,image_width,image_height],[((1-image_display_scale)/2)*screenWidth,picY*screenHeight,((1-image_display_scale)/2)*screenWidth+image_display_width,picY*screenHeight+image_display_width*image_ratio]);
                Screen('FrameRect', window, [255,255,255],[-frameRect_size+((1-image_display_scale)/2)*screenWidth, -frameRect_size+picY*screenHeight, frameRect_size+((1-image_display_scale)/2)*screenWidth+image_display_width, frameRect_size+picY*screenHeight+image_display_width*image_ratio], frameRect_size)
                Screen('DrawLine', window, [255,255,255], (lineX-halfLength-vertical_line_size/2), (lineY), (lineX+halfLength+vertical_line_size/2), (lineY), main_line_size);
                for scalepoint = -3:3
                    Screen('DrawLine', window, [255,255,255], (lineX + scalepoint*divider), (lineY), (lineX + scalepoint*divider), (lineY-vertical_line_height), vertical_line_size);
                end
                Screen('FillRect', window, [255,0,0], [lineX+divider*(currentRating-15)/5-RectWidth,lineY-RectWidth*3,lineX+divider*(currentRating-15)/5+RectWidth,lineY+RectWidth*3]);
                DrawFormattedText(window, double(lowerText), lineX-halfLength-0.04*screenWidth, lineY+0.12*screenHeight, 255)
                DrawFormattedText(window, double(upperText), lineX+halfLength-0.04*screenWidth, lineY+0.12*screenHeight, 255)

                Screen('Flip', window, t - slack);
                try_starttime(k) = GetSecs;
                WaitSecs(2);

                % 初始化计时与变量
                confirmed = false;
                timeoutFlag = 0;
                timeoutShown = false;
                ratingDurationLimit = 8;  % 正常评分时长上限
                feedbackMin = 0.5;        % 超时后的固定反馈时长
                try_startRating(k) = currentRating;
                m = 0;


                %%AI的道德评分
                % 定义图片路径
                if mod(tryconditionList(1,try_blocktime),4)==0
                    try_subADblock(k)=4;
                else
                    try_subADblock(k)=mod(tryconditionList(1,try_blocktime),4);
                end
                imagePath = ['E:\\fMRI代码及其附属文件\\反转小人图片\\yellow' num2str(try_subADblock(k)) '.png'];
                originalImage = imread(imagePath);

                % 计算缩放比例，确保图片适应屏幕宽度或高度
                scaleFactor = min(screenWidth / imageWidth1, screenHeight / imageHeight1);
                scaleFactor = scaleFactor / 2;

                % 调整图片大小
                resizedImage = imresize(originalImage, scaleFactor);

                % 将图片放置在屏幕正中下方
                resizedImageHeight = size(resizedImage, 1);
                resizedImageWidth = size(resizedImage, 2);
                xOffset = (screenWidth - resizedImageWidth) / 2;
                yOffset = screenHeight - resizedImageHeight - 200; % 200像素的底部边距

                % 打印路径，检查是否正确
                disp(imagePath);

                % 检查图片是否存在
                if isempty(originalImage)
                    error('图片无法加载，请检查路径和文件名: %s', imagePath);
                end

                % 在屏幕缓冲区中绘制调整后的图片
                rect = [xOffset, yOffset, xOffset + resizedImageWidth, yOffset + resizedImageHeight];
                Screen('PutImage', window, resizedImage, rect);
                if FTtryrating(k) < 0 || FTtryrating(k) > 30
                    error('try_moralRating1(i) must be between 0 and 30.'); % 确保x在[0,100]之间
                end

                % 定义两个区间
                interval1 = [(max(0, FTtryrating(k) - 9)),(max(0, FTtryrating(k) - 3))];

                interval2 = [(min(30, FTtryrating(k) + 3)), (min(30, FTtryrating(k) + 9))];

                % 检查区间是否合理
                if interval1(1) > interval1(2) || interval2(1) > interval2(2)
                    error('区间定义不合理: interval1 = [%d, %d], interval2 = [%d, %d]', interval1(1), interval1(2), interval2(1), interval2(2));
                end

                % 选择一个区间
                if blockmp(1,i) == 1
                    try_AD_Rating(k) = randi([interval1(1), interval1(2)]);
                    AD_suffered(k)=1;
                else
                    try_AD_Rating(k) = randi([interval2(1), interval2(2)]);
                    AD_suffered(k)=2;
                end
                % 确保结果在[0,30]之间
                try_AD_Rating(k) = max(0, min(30, try_AD_Rating(k)));

                % 调试信息
                disp(['try_moralRating1(i): ', num2str(FTtryrating(k))]);
                disp(['interval1: ', num2str(interval1)]);
                disp(['interval2: ', num2str(interval2)]);
                disp(['try_AI_moralRating(i): ', num2str(try_AD_Rating(k))]);
                ADjittershow(k)=GetSecs;
                Screen('PutImage', window, resizedImage, rect);
                Screen('Flip', window, t - slack);
                WaitSecs(try_AD_jitter(j));


                %% 主评分循环
                try_startA(k) = GetSecs;  % 开始呈现评分界面的时间点
                while ~confirmed
                    [keyIsDown, ~, keyCode] = KbCheck;

                    % 键盘操作
                    if keyIsDown
                        if keyCode(leftKey)
                            currentRating = max(0, currentRating - 1);
                            if try_startB(k) == 0
                                try_startB(k) = GetSecs; % 第一次移动时间
                            end
                            WaitSecs(0.1);
                        elseif keyCode(rightKey)
                            currentRating = min(30, currentRating + 1);
                            if try_startB(k) == 0
                                try_startB(k) = GetSecs;
                            end
                            WaitSecs(0.1);
                        elseif keyCode(makesureKey)
                            try_startC(k) = GetSecs; % 按下B的时间
                            confirmed = true;
                            WaitSecs(0.2);
                        end
                    end

                    % 检查是否超时（超过8秒未按B）
                    if ~timeoutShown && (GetSecs - try_startA(k)) > ratingDurationLimit
                        timeoutFlag = 1;
                        timeoutShown = true;

                        % 显示超时提示1秒
                        overtimeshow(k)=GetSecs;
                        DrawFormattedText(window, double('您已超时，将进入下一回合'), 'center', 'center', [255,0,0]);
                        Screen('Flip', window);
                        WaitSecs(1);
                        confirmed = true;
                        break;
                    end
                    % 绘制评分界面
                    ratingText = [num2str(currentRating)];
                    try_RatingText = [num2str(FTtryrating(k))];
                    if currentRating <=9
                        DrawFormattedText(window, ratingText, lineX-centerCorr1+divider*(currentRating-15)/5, lineY+0.07*screenHeight,[0,255,0]);
                    else
                        DrawFormattedText(window, ratingText, lineX-centerCorr2+divider*(currentRating-15)/5, lineY+0.07*screenHeight, [0,255,0]);
                    end
                    if FTtryrating(k) <=9
                        DrawFormattedText(window, try_RatingText, lineX-centerCorr1+divider*(FTtryrating(k)-15)/5, lineY-0.04*screenHeight, [255,0,0]);
                    else
                        DrawFormattedText(window, try_RatingText, lineX-centerCorr2+divider*(FTtryrating(k)-15)/5, lineY-0.04*screenHeight, [255,0,0]);
                    end

                    % 重新绘制刻度线和文本
                    DrawFormattedText(window, double('您估计对方会在多大程度上觉得该词语所描述的行为是不道德的？'),'center', 0.43*screenHeight, 255);
                    Screen('DrawTexture', window, pic, [0,0,image_width,image_height],[((1-image_display_scale)/2)*screenWidth,picY*screenHeight,((1-image_display_scale)/2)*screenWidth+image_display_width,picY*screenHeight+image_display_width*image_ratio]);
                    Screen('FrameRect', window, [255,255,255],[-frameRect_size+((1-image_display_scale)/2)*screenWidth, -frameRect_size+picY*screenHeight, frameRect_size+((1-image_display_scale)/2)*screenWidth+image_display_width, frameRect_size+picY*screenHeight+image_display_width*image_ratio], frameRect_size)

                    % 画水平主刻度线
                    Screen('DrawLine', window, [255,255,255], (lineX-halfLength-vertical_line_size/2), (lineY), (lineX+halfLength+vertical_line_size/2), (lineY), main_line_size);
                    for scalepoint = -3:3
                        Screen('DrawLine', window, [255,255,255], (lineX + scalepoint*divider), (lineY), (lineX + scalepoint*divider), (lineY-vertical_line_height), vertical_line_size);
                    end
                    DrawFormattedText(window, double(lowerText), lineX-halfLength-0.04*screenWidth, lineY+0.12*screenHeight, 255)
                    DrawFormattedText(window, double(upperText), lineX+halfLength-0.04*screenWidth, lineY+0.12*screenHeight, 255)

                    % 绘制鼠标所在位置的矩形
                    Screen('FillRect', window, [0,255,0], [lineX+divider*(currentRating-15)/5-RectWidth*2,lineY-RectWidth*2,lineX+divider*(currentRating-15)/5+RectWidth*2,lineY+RectWidth*2]);
                    Screen('FillRect', window, [255,0,0], [lineX+divider*(FTtryrating(k)-15)/5-RectWidth,lineY-RectWidth*3,lineX+divider*(FTtryrating(k)-15)/5+RectWidth,lineY+RectWidth*3]);

                    % 重新绘制 AI 的道德评分条
                    AD_rating = try_AD_Rating(k);
                    AD_ratingText = [num2str(AD_rating)];
                    if AD_rating <=9
                        DrawFormattedText(window,AD_ratingText, lineX-centerCorr1+divider*(AD_rating-15)/5, lineY-0.04*screenHeight,  [255,255,0]);
                    else
                        DrawFormattedText(window, AD_ratingText, lineX-centerCorr2+divider*(AD_rating-15)/5, lineY-0.04*screenHeight,  [255,255,0]);
                    end
                    Screen('FillRect', window,  [255,255,0], [lineX+divider*(AD_rating-15)/5-RectWidth,lineY-RectWidth*3,lineX+divider*(AD_rating-15)/5+RectWidth,lineY+RectWidth*3]);
                    Screen('PutImage', window, resizedImage, rect);
                    Screen('Flip', window, t-slack);

                end

                %% 记录时间点与计算结果
                if try_startC(k)~=0
                    try_RT_AC(k) = try_startC(k) - try_startA(k); % 反应时
                else
                    try_RT_AC(k)=0;
                end
                if try_startB(k)~=0
                    try_RT_AB(k) = try_startB(k) - try_startA(k); % 准备移动时长
                else
                    try_RT_AB(k)=0;
                end
                try_Rating(k) = round(currentRating);
                try_isTimeout(k) = timeoutFlag;


                % 若发生过超时提示，记录超时时长及反馈时长
                if timeoutFlag == 1
                    try_feedbackTime(k) = 0;
                else
                    feedbackshow(k)= GetSecs;
                    try_feedbackTime(k) = ratingDurationLimit - try_RT_AC(k);
                    % 显示被试的第二次道德评分反馈界面，同时保留 AI 的道德评分条
                    try_RatingText = [num2str(FTtryrating(k))];
                    if currentRating <=9
                        DrawFormattedText(window, ratingText, lineX-centerCorr1+divider*(currentRating-15)/5, lineY+0.07*screenHeight,[255,0,0]);
                    else
                        DrawFormattedText(window, ratingText, lineX-centerCorr2+divider*(currentRating-15)/5, lineY+0.07*screenHeight, [255,0,0]);
                    end
                    if FTtryrating(k) <=9
                        DrawFormattedText(window, try_RatingText, lineX-centerCorr1+divider*(FTtryrating(k)-15)/5, lineY-0.04*screenHeight, [255,0,0]);
                    else
                        DrawFormattedText(window, try_RatingText, lineX-centerCorr2+divider*(FTtryrating(k)-15)/5, lineY-0.04*screenHeight, [255,0,0]);
                    end
                    DrawFormattedText(window, double('您估计对方会在多大程度上觉得该词语所描述的行为是不道德的？'),'center', 0.43*screenHeight, 255);
                    Screen('DrawTexture', window, pic, [0,0,image_width,image_height],[((1-image_display_scale)/2)*screenWidth,picY*screenHeight,((1-image_display_scale)/2)*screenWidth+image_display_width,picY*screenHeight+image_display_width*image_ratio]);
                    Screen('FrameRect', window, [255,255,255],[-frameRect_size+((1-image_display_scale)/2)*screenWidth, -frameRect_size+picY*screenHeight, frameRect_size+((1-image_display_scale)/2)*screenWidth+image_display_width, frameRect_size+picY*screenHeight+image_display_width*image_ratio], frameRect_size)
                    Screen('DrawLine', window, [255,255,255], (lineX-halfLength-vertical_line_size/2), (lineY), (lineX+halfLength+vertical_line_size/2), (lineY), main_line_size);
                    for scalepoint = -3:3
                        Screen('DrawLine', window, [255,255,255], (lineX + scalepoint*divider), (lineY), (lineX + scalepoint*divider), (lineY-vertical_line_height), vertical_line_size);
                    end
                    DrawFormattedText(window, double(lowerText), lineX-halfLength-0.04*screenWidth, lineY+0.12*screenHeight, 255)
                    DrawFormattedText(window, double(upperText), lineX+halfLength-0.04*screenWidth, lineY+0.12*screenHeight, 255)
                    Screen('FillRect', window, [255,0,0], [lineX+divider*(currentRating-15)/5-RectWidth*2,lineY-RectWidth*2,lineX+divider*(currentRating-15)/5+RectWidth*2,lineY+RectWidth*2]);

                    % 重新绘制 AI 的道德评分条
                    Screen('FillRect', window, [255,0,0], [lineX+divider*(FTtryrating(k)-15)/5-RectWidth,lineY-RectWidth*3,lineX+divider*(FTtryrating(k)-15)/5+RectWidth,lineY+RectWidth*3]);
                    if AD_rating <=9
                        DrawFormattedText(window,AD_ratingText, lineX-centerCorr1+divider*(AD_rating-15)/5, lineY-0.04*screenHeight,  [255,255,0]);
                    else
                        DrawFormattedText(window, AD_ratingText, lineX-centerCorr2+divider*(AD_rating-15)/5, lineY-0.04*screenHeight,  [255,255,0]);
                    end
                    Screen('FillRect', window, [255,255,0], [lineX+divider*(AD_rating-15)/5-RectWidth,lineY-RectWidth*3,lineX+divider*(AD_rating-15)/5+RectWidth,lineY+RectWidth*3]);
                    Screen('PutImage', window, resizedImage, rect);
                    Screen('Flip', window, t-slack);
                    WaitSecs(try_feedbackTime(k));
                end

                %jitter屏
                trailjittershow(k)=GetSecs;
                if mod(k,15)~=0
                    % 绘制白色十字注视点
                    Screen('DrawLines', window, allCoords, lineWidthPix, [255 255 255], [xCenter yCenter]);
                    Screen('Flip', window);
                    if timeoutFlag == 1
                        WaitSecs((jitter_try_shuffled(j)-1));
                    else
                        WaitSecs(jitter_try_shuffled(j));
                    end
                end
                try_block(k)=tryconditionList(1,try_blocktime);
                try_subConditionblock(k)= 1;
                blocksubmergeorder(k)=try_blocktime;
                cueshow(k)=cueshowtime;
                relaxshow(k)=relaxshowtime;
                readyshow(k)=readyshowtime;
            end
        elseif tryconditionList(1,try_blocktime)==9 ||  tryconditionList(1,try_blocktime)==10 ||tryconditionList(1,try_blocktime)==11 ||  tryconditionList(1,try_blocktime)==12 || tryconditionList(1,try_blocktime)==13 ||  tryconditionList(1,try_blocktime)==14 ||tryconditionList(1,try_blocktime)==15 ||  tryconditionList(1,try_blocktime)==16
            %指导语
            cueshowtime=GetSecs;
            if tryconditionList(1,try_blocktime)==9 || tryconditionList(1,try_blocktime)==13
                showInstruction(window, "E:\fMRI代码及其附属文件\指导语\cropped_5.bmp",  slack, [0, 0, 2160, 1440]);
            elseif tryconditionList(1,try_blocktime)==10 ||  tryconditionList(1,try_blocktime)==14
                showInstruction(window, "E:\fMRI代码及其附属文件\指导语\cropped_6.bmp",  slack, [0, 0, 2160, 1440]);
            elseif tryconditionList(1,try_blocktime)==11 || tryconditionList(1,try_blocktime)==15
                showInstruction(window, "E:\fMRI代码及其附属文件\指导语\cropped_7.bmp",  slack, [0, 0, 2160, 1440]);
            elseif tryconditionList(1,try_blocktime)==12 || tryconditionList(1,try_blocktime)==16
                showInstruction(window, "E:\fMRI代码及其附属文件\指导语\cropped_8.bmp",  slack, [0, 0, 2160, 1440]);
            end
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
            for i = 1:food_try_numRows
                k=15*(tryconditionList(1,try_blocktime)-1)+i;
                j=15*(try_blocktime-1)+i;
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
                lowerText = '完全没有不喜欢';
                upperText = '非常不喜欢';
                lineLength = 0.5 * screenWidth;
                halfLength = lineLength / 2;
                divider = lineLength / 6;
                lineX = XCenter;
                lineY = 0.55 * screenHeight;
                y_mouse = lineY;

                currentRating=FTtryrating(k);
                try_RatingText = [num2str(FTtryrating(k))];
                if FTtryrating(k) <=9
                    DrawFormattedText(window, try_RatingText, lineX-centerCorr1+divider*(FTtryrating(k)-15)/5, lineY-0.04*screenHeight, [255,0,0]);
                else
                    DrawFormattedText(window, try_RatingText, lineX-centerCorr2+divider*(FTtryrating(k)-15)/5, lineY-0.04*screenHeight, [255,0,0]);
                end
                DrawFormattedText(window, double('您的食物初次估计值为'),'center', 0.43*screenHeight, 255);
                Screen('DrawTexture', window, pic, [0,0,image_width,image_height],[((1-image_display_scale)/2)*screenWidth,picY*screenHeight,((1-image_display_scale)/2)*screenWidth+image_display_width,picY*screenHeight+image_display_width*image_ratio]);
                Screen('FrameRect', window, [255,255,255],[-frameRect_size+((1-image_display_scale)/2)*screenWidth, -frameRect_size+picY*screenHeight, frameRect_size+((1-image_display_scale)/2)*screenWidth+image_display_width, frameRect_size+picY*screenHeight+image_display_width*image_ratio], frameRect_size)
                Screen('DrawLine', window, [255,255,255], (lineX-halfLength-vertical_line_size/2), (lineY), (lineX+halfLength+vertical_line_size/2), (lineY), main_line_size);
                for scalepoint = -3:3
                    Screen('DrawLine', window, [255,255,255], (lineX + scalepoint*divider), (lineY), (lineX + scalepoint*divider), (lineY-vertical_line_height), vertical_line_size);
                end
                Screen('FillRect', window, [255,0,0], [lineX+divider*(currentRating-15)/5-RectWidth,lineY-RectWidth*3,lineX+divider*(currentRating-15)/5+RectWidth,lineY+RectWidth*3]);
                DrawFormattedText(window, double(lowerText), lineX-halfLength-0.04*screenWidth, lineY+0.12*screenHeight, 255)
                DrawFormattedText(window, double(upperText), lineX+halfLength-0.04*screenWidth, lineY+0.12*screenHeight, 255)

                Screen('Flip', window, t - slack);
                try_starttime(k) = GetSecs;
                WaitSecs(2);

                % 初始化计时与变量
                confirmed = false;
                timeoutFlag = 0;
                timeoutShown = false;
                ratingDurationLimit = 8;  % 正常评分时长上限
                feedbackMin = 0.5;        % 超时后的固定反馈时长
                try_startRating(k) = currentRating;
                m = 0;


                %%AI的道德评分
                % 定义图片路径
                if mod(tryconditionList(1,try_blocktime),4)==0
                    try_subADblock(k)=4;
                else
                    try_subADblock(k)=mod(tryconditionList(1,try_blocktime),4);
                end
                imagePath = ['E:\\fMRI代码及其附属文件\\反转小人图片\\yellow' num2str(try_subADblock(k)) '.png'];
                originalImage = imread(imagePath);

                % 计算缩放比例，确保图片适应屏幕宽度或高度
                scaleFactor = min(screenWidth / imageWidth1, screenHeight / imageHeight1);
                scaleFactor = scaleFactor / 2;

                % 调整图片大小
                resizedImage = imresize(originalImage, scaleFactor);

                % 将图片放置在屏幕正中下方
                resizedImageHeight = size(resizedImage, 1);
                resizedImageWidth = size(resizedImage, 2);
                xOffset = (screenWidth - resizedImageWidth) / 2;
                yOffset = screenHeight - resizedImageHeight - 200; % 200像素的底部边距

                % 打印路径，检查是否正确
                disp(imagePath);

                % 检查图片是否存在
                if isempty(originalImage)
                    error('图片无法加载，请检查路径和文件名: %s', imagePath);
                end

                % 在屏幕缓冲区中绘制调整后的图片
                rect = [xOffset, yOffset, xOffset + resizedImageWidth, yOffset + resizedImageHeight];
                Screen('PutImage', window, resizedImage, rect);
                if FTtryrating(k) < 0 || FTtryrating(k) > 30
                    error('try_moralRating1(i) must be between 0 and 30.'); % 确保x在[0,100]之间
                end

                % 定义两个区间
                interval1 = [(max(0, FTtryrating(k) - 9)),(max(0, FTtryrating(k) - 3))];

                interval2 = [(min(30, FTtryrating(k) + 3)), (min(30, FTtryrating(k) + 9))];

                % 检查区间是否合理
                if interval1(1) > interval1(2) || interval2(1) > interval2(2)
                    error('区间定义不合理: interval1 = [%d, %d], interval2 = [%d, %d]', interval1(1), interval1(2), interval2(1), interval2(2));
                end

                %选择一个区间
                if blockmp(1,i) == 1
                    try_AD_Rating(k) = randi([interval1(1), interval1(2)]);
                    AD_suffered(k)=1;
                else
                    try_AD_Rating(k) = randi([interval2(1), interval2(2)]);
                    AD_suffered(k)=2;
                end


                % 确保结果在[0,30]之间
                try_AD_Rating(k) = max(0, min(30, try_AD_Rating(k)));

                % 调试信息
                disp(['try_moralRating1(i): ', num2str(FTtryrating(k))]);
                disp(['interval1: ', num2str(interval1)]);
                disp(['interval2: ', num2str(interval2)]);
                disp(['try_AI_moralRating(i): ', num2str(try_AD_Rating(k))]);
                ADjittershow(k)=GetSecs;
                Screen('PutImage', window, resizedImage, rect);
                Screen('Flip', window, t - slack);
                WaitSecs(try_AD_jitter(j));


                %% 主评分循环
                try_startA(k) = GetSecs;  % 开始呈现评分界面的时间点
                while ~confirmed
                    [keyIsDown, ~, keyCode] = KbCheck;

                    % 检查是否超时（超过8秒未按B）
                    if ~timeoutShown && (GetSecs - try_startA(k)) > ratingDurationLimit
                        timeoutFlag = 1;
                        timeoutShown = true;

                        % 显示超时提示1秒
                        overtimeshow(k)=GetSecs;
                        DrawFormattedText(window, double('您已超时，将进入下一回合'), 'center', 'center', [255,0,0]);
                        Screen('Flip', window);
                        WaitSecs(1);
                        confirmed = true;
                        break;
                    end


                    % 键盘操作
                    if keyIsDown
                        if keyCode(leftKey)
                            currentRating = max(0, currentRating - 1);
                            if try_startB(k) == 0
                                try_startB(k) = GetSecs; % 第一次移动时间
                            end
                            WaitSecs(0.1);
                        elseif keyCode(rightKey)
                            currentRating = min(30, currentRating + 1);
                            if try_startB(k) == 0
                                try_startB(k) = GetSecs;
                            end
                            WaitSecs(0.1);
                        elseif keyCode(makesureKey)
                            try_startC(k) = GetSecs; % 按下B的时间
                            confirmed = true;
                            WaitSecs(0.2);
                        end
                    end

                    % 绘制评分界面
                    ratingText = [num2str(currentRating)];
                    try_RatingText = [num2str(FTtryrating(k))];
                    if currentRating <=9
                        DrawFormattedText(window, ratingText, lineX-centerCorr1+divider*(currentRating-15)/5, lineY+0.07*screenHeight,[0,255,0]);
                    else
                        DrawFormattedText(window, ratingText, lineX-centerCorr2+divider*(currentRating-15)/5, lineY+0.07*screenHeight, [0,255,0]);
                    end
                    if FTtryrating(k) <=9
                        DrawFormattedText(window, try_RatingText, lineX-centerCorr1+divider*(FTtryrating(k)-15)/5, lineY-0.04*screenHeight, [255,0,0]);
                    else
                        DrawFormattedText(window, try_RatingText, lineX-centerCorr2+divider*(FTtryrating(k)-15)/5, lineY-0.04*screenHeight, [255,0,0]);
                    end

                    % 重新绘制刻度线和文本
                    DrawFormattedText(window, double('您估计对方会在多大程度上不喜欢该词语所描述的食物？'),'center', 0.43*screenHeight, 255);
                    Screen('DrawTexture', window, pic, [0,0,image_width,image_height],[((1-image_display_scale)/2)*screenWidth,picY*screenHeight,((1-image_display_scale)/2)*screenWidth+image_display_width,picY*screenHeight+image_display_width*image_ratio]);
                    Screen('FrameRect', window, [255,255,255],[-frameRect_size+((1-image_display_scale)/2)*screenWidth, -frameRect_size+picY*screenHeight, frameRect_size+((1-image_display_scale)/2)*screenWidth+image_display_width, frameRect_size+picY*screenHeight+image_display_width*image_ratio], frameRect_size)

                    % 画水平主刻度线
                    Screen('DrawLine', window, [255,255,255], (lineX-halfLength-vertical_line_size/2), (lineY), (lineX+halfLength+vertical_line_size/2), (lineY), main_line_size);
                    for scalepoint = -3:3
                        Screen('DrawLine', window, [255,255,255], (lineX + scalepoint*divider), (lineY), (lineX + scalepoint*divider), (lineY-vertical_line_height), vertical_line_size);
                    end
                    DrawFormattedText(window, double(lowerText), lineX-halfLength-0.04*screenWidth, lineY+0.12*screenHeight, 255)
                    DrawFormattedText(window, double(upperText), lineX+halfLength-0.04*screenWidth, lineY+0.12*screenHeight, 255)

                    % 绘制鼠标所在位置的矩形
                    Screen('FillRect', window, [0,255,0], [lineX+divider*(currentRating-15)/5-RectWidth*2,lineY-RectWidth*2,lineX+divider*(currentRating-15)/5+RectWidth*2,lineY+RectWidth*2]);
                    Screen('FillRect', window, [255,0,0], [lineX+divider*(FTtryrating(k)-15)/5-RectWidth,lineY-RectWidth*3,lineX+divider*(FTtryrating(k)-15)/5+RectWidth,lineY+RectWidth*3]);

                    % 重新绘制 AI 的道德评分条
                    AD_rating = try_AD_Rating(k);
                    AD_ratingText = [num2str(AD_rating)];
                    if AD_rating <=9
                        DrawFormattedText(window,AD_ratingText, lineX-centerCorr1+divider*(AD_rating-15)/5, lineY-0.04*screenHeight,  [255,255,0]);
                    else
                        DrawFormattedText(window, AD_ratingText, lineX-centerCorr2+divider*(AD_rating-15)/5, lineY-0.04*screenHeight,  [255,255,0]);
                    end
                    Screen('FillRect', window,  [255,255,0], [lineX+divider*(AD_rating-15)/5-RectWidth,lineY-RectWidth*3,lineX+divider*(AD_rating-15)/5+RectWidth,lineY+RectWidth*3]);
                    Screen('PutImage', window, resizedImage, rect);
                    Screen('Flip', window, t-slack);
                end

                %% 记录时间点与计算结果
                if try_startC(k)~=0
                    try_RT_AC(k) = try_startC(k) - try_startA(k); % 反应时
                else
                    try_RT_AC(k)=0;
                end
                if try_startB(k)~=0
                    try_RT_AB(k) = try_startB(k) - try_startA(k); % 准备移动时长
                else
                    try_RT_AB(k)=0;
                end
                try_Rating(k) = round(currentRating);
                try_isTimeout(k) = timeoutFlag;


                % 反馈屏
                if timeoutFlag == 1
                    try_feedbackTime(k) = 0;
                else
                    feedbackshow(k)= GetSecs;
                    try_feedbackTime(k) = ratingDurationLimit - try_RT_AC(k);
                    try_RatingText = [num2str(FTtryrating(k))];
                    if currentRating <=9
                        DrawFormattedText(window, ratingText, lineX-centerCorr1+divider*(currentRating-15)/5, lineY+0.07*screenHeight,[255,0,0]);
                    else
                        DrawFormattedText(window, ratingText, lineX-centerCorr2+divider*(currentRating-15)/5, lineY+0.07*screenHeight, [255,0,0]);
                    end
                    if FTtryrating(k) <=9
                        DrawFormattedText(window, try_RatingText, lineX-centerCorr1+divider*(FTtryrating(k)-15)/5, lineY-0.04*screenHeight, [255,0,0]);
                    else
                        DrawFormattedText(window, try_RatingText, lineX-centerCorr2+divider*(FTtryrating(k)-15)/5, lineY-0.04*screenHeight, [255,0,0]);
                    end
                    DrawFormattedText(window, double('您估计对方会在多大程度上不喜欢该词语所描述的食物？'),'center', 0.43*screenHeight, 255);
                    Screen('DrawTexture', window, pic, [0,0,image_width,image_height],[((1-image_display_scale)/2)*screenWidth,picY*screenHeight,((1-image_display_scale)/2)*screenWidth+image_display_width,picY*screenHeight+image_display_width*image_ratio]);
                    Screen('FrameRect', window, [255,255,255],[-frameRect_size+((1-image_display_scale)/2)*screenWidth, -frameRect_size+picY*screenHeight, frameRect_size+((1-image_display_scale)/2)*screenWidth+image_display_width, frameRect_size+picY*screenHeight+image_display_width*image_ratio], frameRect_size)
                    Screen('DrawLine', window, [255,255,255], (lineX-halfLength-vertical_line_size/2), (lineY), (lineX+halfLength+vertical_line_size/2), (lineY), main_line_size);
                    for scalepoint = -3:3
                        Screen('DrawLine', window, [255,255,255], (lineX + scalepoint*divider), (lineY), (lineX + scalepoint*divider), (lineY-vertical_line_height), vertical_line_size);
                    end
                    DrawFormattedText(window, double(lowerText), lineX-halfLength-0.04*screenWidth, lineY+0.12*screenHeight, 255)
                    DrawFormattedText(window, double(upperText), lineX+halfLength-0.04*screenWidth, lineY+0.12*screenHeight, 255)
                    Screen('FillRect', window, [255,0,0], [lineX+divider*(currentRating-15)/5-RectWidth*2,lineY-RectWidth*2,lineX+divider*(currentRating-15)/5+RectWidth*2,lineY+RectWidth*2]);

                    % 重新绘制 AI 的道德评分条
                    Screen('FillRect', window, [255,0,0], [lineX+divider*(FTtryrating(k)-15)/5-RectWidth,lineY-RectWidth*3,lineX+divider*(FTtryrating(k)-15)/5+RectWidth,lineY+RectWidth*3]);
                    if AD_rating <=9
                        DrawFormattedText(window,AD_ratingText, lineX-centerCorr1+divider*(AD_rating-15)/5, lineY-0.04*screenHeight,  [255,255,0]);
                    else
                        DrawFormattedText(window, AD_ratingText, lineX-centerCorr2+divider*(AD_rating-15)/5, lineY-0.04*screenHeight,  [255,255,0]);
                    end
                    Screen('FillRect', window, [255,255,0], [lineX+divider*(AD_rating-15)/5-RectWidth,lineY-RectWidth*3,lineX+divider*(AD_rating-15)/5+RectWidth,lineY+RectWidth*3]);
                    Screen('PutImage', window, resizedImage, rect);
                    Screen('Flip', window, t-slack);
                    WaitSecs(try_feedbackTime(k));
                end

                %jitter屏
                trailjittershow(k)=GetSecs;
                if mod(k,15)~=0
                    % 绘制白色十字注视点
                    Screen('DrawLines', window, allCoords, lineWidthPix, [255 255 255], [xCenter yCenter]);
                    Screen('Flip', window);
                    if timeoutFlag == 1
                        WaitSecs((jitter_try_shuffled(j)-1));
                    else
                        WaitSecs(jitter_try_shuffled(j));
                    end
                end
                try_block(k)=tryconditionList(1,try_blocktime);
                try_subConditionblock(k)= 2;
                blocksubmergeorder(k)=try_blocktime;
                cueshow(k)=cueshowtime;
                relaxshow(k)=relaxshowtime;
                readyshow(k)=readyshowtime;
            end
        end
        if mod(try_blocktime,2)==0 && try_blocktime<16
            blockovertime=GetSecs;
            relaxshowtime= GetSecs;
            DrawFormattedText(window, double('您可以短暂休息一会儿'), 'center', 'center', [255,255,255]);
            Screen('Flip', window);
            triggerTime = wait_for_trigger_advanced();
            if triggerTime>blockovertime
                readyshowtime= GetSecs;
                DrawFormattedText(window, double('将继续任务，请做好准备'), 'center', 'center', 255);
                Screen('Flip', window);
                WaitSecs(2);
            end
        end
        try_blocktime=try_blocktime+1;
    end

    ending_Key = KbName('y');
    while 1
        DrawFormattedText(window, double('您已完成评分阶段的全部任务，请联系实验人员'),'center', 'center', [255,255,255]);
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
    % 获取 subInfo 的长度
    try_subinfo_length = length(try_subinfo);

    % 如果 subInfo 的长度小于 try_trialnum，则复制最后一行
    if try_subinfo_length < try_trialnum
        % 计算需要复制的次数
        num_repeats = try_trialnum - try_subinfo_length;

        % 复制最后一行
        try_subinfo = [try_subinfo; repmat(try_subinfo(end, :), num_repeats, 1)];
    end


    if ~all([length(try_picID), length(readyshow),length(relaxshow),length(overtimeshow), length(try_starttime),length(try_startA), length(feedbackshow),length(try_startB),length(try_startC),length(try_startRating),length(try_Rating), length(try_RT_AC), length(try_RT_AB),length(try_feedbackTime),length(try_isTimeout),length(try_AD_jitter), length(try_AD_Rating),length(jitter_try_shuffled),length(cueshow),length(ADjittershow),length(trailjittershow),length(blocksubmergeorder),length(try_subADblock),length(try_subConditionblock),length(try_block)] == try_trialnum)
        error('数组长度不一致，请检查数据收集部分的代码');
    end
    rating_try_data = [try_subinfo, try_picID,readyshow, cueshow,try_starttime, ADjittershow,try_startA,try_startB,try_startC,feedbackshow,overtimeshow,trailjittershow,relaxshow,try_startRating, try_Rating,try_AD_Rating,try_RT_AC,try_RT_AB,try_feedbackTime,try_isTimeout,try_AD_jitter,AD_suffered,jitter_try_shuffled,blocksubmergeorder,try_subADblock,try_subConditionblock,try_block];
    variableNames = {'subInfo', 'picID','readyshowtime','cueshowtime','showwordtime','ADjittershowtime','ScreenStartTime', 'MouseMoveTime', 'ClickTime', 'feebackshowtime','overtimeshow','trailjittershow','relaxshow','randomtry_startRating','Rating','AD_Rating','reacttime','premovetime','feedbacktime','istimeout','ADjitter','ADPlus','jitter_try_shuffled','blockorder','try_subADblock','try_subConditionblock','try_block'}; % 列标题
    rating_try_result = table(try_subinfo, try_picID, readyshow,cueshow,try_starttime, ADjittershow,try_startA,try_startB,try_startC,feedbackshow,overtimeshow,trailjittershow,relaxshow,try_startRating, try_Rating,try_AD_Rating,try_RT_AC,try_RT_AB,try_feedbackTime,try_isTimeout,try_AD_jitter,AD_suffered,jitter_try_shuffled,blocksubmergeorder,try_subADblock,try_subConditionblock,try_block,'VariableNames', variableNames);

    subnum = subinfo{1};  % 从 cell 里取编号（可能是数字或字符串）
    baseDir = "E:\fMRI代码及其附属文件\实验数据\second_trydata\";

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
    WaitSecs(4);
    WaitSecs(slack);  % 控制刷新频率
end
end

function triggerTime = wait_for_trigger_advanced()
persistent lastTriggerTime
if isempty(lastTriggerTime)
    lastTriggerTime = 0;
end

fprintf('等待trigger...\n');

while true
    [keyIsDown, secs, keyCode] = KbCheck;

    if keyIsDown && keyCode(KbName('s'))
        % 防重复：确保至少间隔0.5秒才认为是新trigger
        if secs > lastTriggerTime + 0.5
            triggerTime = secs;
            lastTriggerTime = secs;
            fprintf('收到trigger！时间: %.6f\n', triggerTime);
            break;
        end
    end

    WaitSecs(0.001);
end
end