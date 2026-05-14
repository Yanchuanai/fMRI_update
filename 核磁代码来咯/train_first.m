clear;clc;
path=pwd;
PsychDefaultSetup(0);

try    %% 被试信息
    subID=17;
    %MATLAB自带inputdlg函数，其作用是打开一个收集用户输入的对话框，并将输入的信息赋给一个变量
    % inputdlg中，prompt是对需要输入的信息的提示，title就是对话框标题，“[1, 50]”是输入框的维度，表示对话框的高度为一个字符，长度为50个字符。
    prompt = {'Subject Number','Gender[1 = male, 2 = female]','Age','Handeness[1 = left, 2 = right]', 'Name[initials]'};
    title = 'Exp infor'; % The title of the dialog boxy
    default_input = {'999','999','999','999','999'}; % Default input value(s)
    subinfo = inputdlg(prompt,title,[1, 50],default_input)';

    %% 读入相关材料
    moralR_train_material_data = readmatrix ("E:\fMRI代码及其附属文件\核磁代码来咯\代码中用到的Excel文件\rating_train.xlsx");
    moralR_train_numRows = moralR_train_material_data(:)';


    %% 实验信息配置
    train_trialnum=8;

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
    train_picID = zeros(train_trialnum, 1);
    train_subinfo = zeros(train_trialnum,1);
    train_pic= zeros(train_trialnum,1);
    train_block= zeros(train_trialnum,1);
    train_subADblock= zeros(train_trialnum,1);
    train_subConditionblock= zeros(train_trialnum,1);

    train_starttime = zeros(train_trialnum, 1);
    train_startA = zeros(train_trialnum, 1);
    train_startB = zeros(train_trialnum, 1);
    train_startC = zeros(train_trialnum, 1);
    train_startRating = zeros(train_trialnum,1);
    train_Rating = zeros(train_trialnum,1);
    train_RT_AC = zeros(train_trialnum, 1);
    train_RT_AB = zeros(train_trialnum, 1);


    train_confidencestartA = zeros(train_trialnum, 1);
    train_confidencestartB = zeros(train_trialnum, 1);
    train_confidencestartC = zeros(train_trialnum, 1);
    train_confidencestartRating = zeros(train_trialnum,1);
    train_confidenceRating = zeros(train_trialnum,1);
    train_confidenceRT_AC = zeros(train_trialnum, 1);
    train_confidenceRT_AB = zeros(train_trialnum, 1);

    for i = 1: length(subinfo)
        train_subinfo(i) = str2double(subinfo{i});
    end
    % 初始化 train_pic 为单元数组
    train_pic = cell(train_trialnum, 1);

    % 练习试次图片导入
    for i = 1:train_trialnum
        train_image_ID = moralR_train_numRows(1,i);
        filePath = ['E:\fMRI代码及其附属文件\练习试次图片11\' num2str(train_image_ID) '.bmp'];
        disp(filePath);  % 打印路径，检查是否正确
        if ~exist(filePath, 'file')
            error('文件不存在，请检查路径和文件名: %s', filePath);
        end
        train_pic{i} = imread(filePath);
    end

    train_blocktime=1;
    trainconditionList = [2 3 5 8];
    % 获取数列的长度
    n = length(trainconditionList);
    % 生成随机索引
    randomIndices = randperm(n);
    % 根据随机索引重新排列数列
    random_trainconditionList = trainconditionList(randomIndices);
    filename ="E:\fMRI代码及其附属文件\临时数据存放点\train_order.xlsx";
    if isfile(filename)
        % 读取已有数据
        existingData = readmatrix(filename);
    else
        % 如果文件不存在，则创建一个空矩阵
        existingData = [];
    end

    % 确保行数够
    if size(existingData,1) < subID
        existingData(subID, 1:length(random_trainconditionList)) = random_trainconditionList;
    else
        existingData(subID, :) = random_trainconditionList;
    end
    % 写入 Excel
    writematrix(existingData, filename);

    showexperimentInstruction(window, "E:\fMRI代码及其附属文件\指导语\croppedinstru_1.bmp",  slack, [0, 0, 2160, 1440]);
    % 暂停 0.5 秒
    WaitSecs(0.5);
    while train_blocktime<=4
        if random_trainconditionList(1,train_blocktime)==2 ||  random_trainconditionList(1,train_blocktime)==3
            showInstruction(window, "E:\fMRI代码及其附属文件\指导语\cropped_9.bmp",  slack, [0, 0, 2160, 1440]);

            %显示练习回合指导语

            for i = 1:2
                if random_trainconditionList(1,train_blocktime)==2
                    k=random_trainconditionList(1,train_blocktime)-2+i;
                else
                    k=random_trainconditionList(1,train_blocktime)-1+i;
                end
                train_image_ID = moralR_train_material_data(k);
                filePath = ['E:\fMRI代码及其附属文件\练习试次图片11\' num2str(train_image_ID) '.bmp'];
                disp(filePath);  % 打印路径，检查是否正确
                if ~exist(filePath, 'file')
                    error('文件不存在，请检查路径和文件名: %s', filePath);
                end
                pic_image = imread(filePath);
                train_picID(k) = train_image_ID;
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
                train_starttime(k) = GetSecs;
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
                train_startA(k) = GetSecs;  % 开始呈现评分界面的时间点
                confirmed = false;
                currentRating = (randi(31)-1);
                train_startRating(k) = currentRating;
                Screen('Flip', window, t - slack);

                %% 主评分循环
                while ~confirmed
                    [keyIsDown, ~, keyCode] = KbCheck;
                    % 键盘操作
                    if keyIsDown
                        if keyCode(KbName('j')) || keyCode(KbName('J'))
                            currentRating = max(0, currentRating - 1);
                            if train_startB(k) == 0
                                train_startB(k) = GetSecs; % 第一次移动时间
                            end
                            WaitSecs(0.1);
                        elseif keyCode(KbName('k')) || keyCode(KbName('K'))
                            currentRating = min(30, currentRating + 1);
                            if train_startB(k) == 0
                                train_startB(k) = GetSecs;
                            end
                            WaitSecs(0.1);
                        elseif keyCode(KbName('f')) || keyCode(KbName('F'))
                            train_startC(k) = GetSecs; % 按下B的时间
                            confirmed = true;
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
                train_RT_AC(k) = train_startC(k) - train_startA(k); % 反应时
                if train_startB(k)~=0
                    train_RT_AB(k) = train_startB(k) - train_startA(k); % 准备移动时长
                else
                    train_RT_AB(k)=0;
                end
                train_Rating(k) = round(currentRating);

                %% 鼠标反馈屏
                if currentRating <=9
                    DrawFormattedText(window, ratingText, lineX-centerCorr1+divider*(currentRating-15)/5, lineY-0.04*screenHeight, [255,0,0]);
                else
                    DrawFormattedText(window, ratingText, lineX-centerCorr2+divider*(currentRating-15)/5, lineY-0.04*screenHeight, [255,0,0]);
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
                train_confidencestartA(k) = GetSecs;  % 开始呈现评分界面的时间点
                confirmed = false;
                currentRating = (randi(31)-1);
                train_confidencestartRating(k) = currentRating;
                m = 0;
                Screen('Flip', window, t - slack);

                %% 主评分循环
                while ~confirmed
                    [keyIsDown, ~, keyCode] = KbCheck;
                    if keyIsDown
                        if keyCode(KbName('j')) || keyCode(KbName('J'))
                            currentRating = max(0, currentRating - 1);
                            if train_confidencestartB(k) == 0
                                train_confidencestartB(k) = GetSecs; % 第一次移动时间
                            end
                            WaitSecs(0.1);
                        elseif keyCode(KbName('k')) || keyCode(KbName('K'))
                            currentRating = min(30, currentRating + 1);
                            if train_confidencestartB(k) == 0
                                train_confidencestartB(k) = GetSecs;
                            end
                            WaitSecs(0.1);
                        elseif keyCode(KbName('f')) || keyCode(KbName('F'))
                            train_confidencestartC(k) = GetSecs; % 按下B的时间
                            confirmed = true;
                            WaitSecs(0.2);
                        end
                    end

                    % 绘制评分界面
                    ratingText = [num2str(currentRating) ];
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
                train_confidenceRT_AC(k) = train_confidencestartC(k) - train_confidencestartA(k); % 反应时
                if train_confidencestartB(k)~=0
                    train_confidenceRT_AB(k) = train_confidencestartB(k) - train_confidencestartA(k); % 准备移动时长
                else
                    train_confidenceRT_AB(k)=0;
                end
                train_confidenceRating(k) = round(currentRating);


                %% 鼠标反馈屏
                if currentRating <=9
                    DrawFormattedText(window, ratingText, lineX-centerCorr1+divider*(currentRating-15)/5, lineY-0.04*screenHeight, [255,0,0]);
                else
                    DrawFormattedText(window, ratingText, lineX-centerCorr2+divider*(currentRating-15)/5, lineY-0.04*screenHeight, [255,0,0]);
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
                if random_trainconditionList(1,train_blocktime)==2
                    train_block(k)=2;
                    train_subADblock(k)= 2;
                    train_subConditionblock(k)= 1;
                else
                    train_block(k)=3;
                    train_subADblock(k)= 1;
                    train_subConditionblock(k)= 1;
                end
            end
        else
            showInstruction(window, "E:\fMRI代码及其附属文件\指导语\cropped_10.bmp",  slack, [0, 0, 2160, 1440]);
            for i = 1:2
                if random_trainconditionList(1,train_blocktime)==5
                    k=random_trainconditionList(1,train_blocktime)-1+i;
                else
                    k=random_trainconditionList(1,train_blocktime)-2+i;
                end
                train_image_ID = moralR_train_material_data(k);
                filePath = ['E:\fMRI代码及其附属文件\练习试次图片11\' num2str(train_image_ID) '.bmp'];
                disp(filePath);  % 打印路径，检查是否正确
                if ~exist(filePath, 'file')
                    error('文件不存在，请检查路径和文件名: %s', filePath);
                end
                pic_image = imread(filePath);
                train_picID(k) = train_image_ID;
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
                train_starttime(k) = GetSecs;
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
                train_startA(k) = GetSecs;  % 开始呈现评分界面的时间点
                confirmed = false;
                currentRating = (randi(31)-1);
                train_startRating(k) = currentRating;
                Screen('Flip', window, t - slack);

                %% 主评分循环
                while ~confirmed
                    [keyIsDown, ~, keyCode] = KbCheck;
                    % 键盘操作
                    if keyIsDown
                        if keyCode(KbName('j')) || keyCode(KbName('J'))
                            currentRating = max(0, currentRating - 1);
                            if train_startB(k) == 0
                                train_startB(k) = GetSecs; % 第一次移动时间
                            end
                            WaitSecs(0.1);
                        elseif keyCode(KbName('k')) || keyCode(KbName('K'))
                            currentRating = min(30, currentRating + 1);
                            if train_startB(k) == 0
                                train_startB(k) = GetSecs;
                            end
                            WaitSecs(0.1);
                        elseif keyCode(KbName('f')) || keyCode(KbName('F'))
                            train_startC(k) = GetSecs; % 按下B的时间
                            confirmed = true;
                            WaitSecs(0.2);
                        end
                    end

                    % 绘制评分界面
                    ratingText = [num2str(currentRating) ];
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
                train_RT_AC(k) = train_startC(k) - train_startA(k); % 反应时
                if train_startB(k)~=0
                    train_RT_AB(k) = train_startB(k) - train_startA(k); % 准备移动时长
                else
                    train_RT_AB(k)=0;
                end
                train_Rating(k) = round(currentRating);

                %% 鼠标反馈屏
                if currentRating <=9
                    DrawFormattedText(window, ratingText, lineX-centerCorr1+divider*(currentRating-15)/5, lineY-0.04*screenHeight, [255,0,0]);
                else
                    DrawFormattedText(window, ratingText, lineX-centerCorr2+divider*(currentRating-15)/5, lineY-0.04*screenHeight, [255,0,0]);
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
                train_confidencestartA(k) = GetSecs;  % 开始呈现评分界面的时间点
                confirmed = false;
                currentRating = (randi(31)-1);
                train_confidencestartRating(k) = currentRating;
                m = 0;
                Screen('Flip', window, t - slack);

                %% 主评分循环
                while ~confirmed
                    [keyIsDown, ~, keyCode] = KbCheck;
                    if keyIsDown
                        if keyCode(KbName('j')) || keyCode(KbName('J'))
                            currentRating = max(0, currentRating - 1);
                            if train_confidencestartB(k) == 0
                                train_confidencestartB(k) = GetSecs; % 第一次移动时间
                            end
                            WaitSecs(0.1);
                        elseif keyCode(KbName('k')) || keyCode(KbName('K'))
                            currentRating = min(30, currentRating + 1);
                            if train_confidencestartB(k) == 0
                                train_confidencestartB(k) = GetSecs;
                            end
                            WaitSecs(0.1);
                        elseif keyCode(KbName('f')) || keyCode(KbName('F'))
                            train_confidencestartC(k) = GetSecs; % 按下B的时间
                            confirmed = true;
                            WaitSecs(0.2);
                        end
                    end

                    % 绘制评分界面
                    ratingText = [num2str(currentRating) ];
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
                train_confidenceRT_AC(k) = train_confidencestartC(k) - train_confidencestartA(k); % 反应时
                if train_confidencestartB(k)~=0
                    train_confidenceRT_AB(k) = train_confidencestartB(k) - train_confidencestartA(k); % 准备移动时长
                else
                    train_confidenceRT_AB(k)=0;
                end
                train_confidenceRating(k) = round(currentRating);


                %% 鼠标反馈屏
                if currentRating <=9
                    DrawFormattedText(window, ratingText, lineX-centerCorr1+divider*(currentRating-15)/5, lineY-0.04*screenHeight, [255,0,0]);
                else
                    DrawFormattedText(window, ratingText, lineX-centerCorr2+divider*(currentRating-15)/5, lineY-0.04*screenHeight, [255,0,0]);
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
                if random_trainconditionList(1,train_blocktime)==5
                    train_block(k)=5;
                    train_subADblock(k)= 1;
                    train_subConditionblock(k)= 2;
                else
                    train_block(k)=8;
                    train_subADblock(k)= 2;
                    train_subConditionblock(k)= 2;
                end
            end
        end
        train_blocktime=train_blocktime+1;
    end



    %显示练习回合指导语
    showexperimentInstruction(window, "E:\fMRI代码及其附属文件\指导语\croppedinstru_3.bmp",  slack, [0, 0, 2160, 1440]);
    % 暂停 0.5 秒
    WaitSecs(0.5);


    Screen('CloseAll');
    Screen('Close');
    fclose('all');
    %% 数据存放

    % 3. 保存到 Excel（没有表头）
    filename = "E:\fMRI代码及其附属文件\临时数据存放点\FTtrainrating.xlsx";
    writematrix(train_Rating, filename, 'WriteMode', 'overwrite');

    % 获取 subInfo 的长度
    train_subinfo_length = length(train_subinfo);

    % 如果 subInfo 的长度小于 train_trialnum，则复制最后一行
    if train_subinfo_length < train_trialnum
        % 计算需要复制的次数
        num_repeats = train_trialnum - train_subinfo_length;

        % 复制最后一行
        train_subinfo = [train_subinfo; repmat(train_subinfo(end, :), num_repeats, 1)];
    end


    if ~all([length(train_picID),  length(train_starttime),length(train_startA), length(train_startB),length(train_startC),length(train_startRating),length(train_Rating), length(train_RT_AC), length(train_RT_AB), length(train_confidencestartA), length(train_confidencestartB),length(train_confidencestartC),length(train_confidencestartRating), length(train_confidenceRating), length(train_confidenceRT_AC), length(train_confidenceRT_AB),length(train_subADblock),length(train_subConditionblock),length(train_block)] == train_trialnum)
        error('数组长度不一致，请检查数据收集部分的代码');
    end
    rating_train_data = [train_subinfo, train_picID, train_starttime, train_startA,train_startB,train_startC,train_startRating, train_Rating,train_RT_AC,train_RT_AB,train_confidencestartA,train_confidencestartB,train_confidencestartC,train_confidencestartRating,train_confidenceRating,train_confidenceRT_AC,train_confidenceRT_AB,train_subADblock,train_subConditionblock,train_block];
    variableNames = {'subInfo', 'picID','showwordtime','ScreenStartTime', 'MouseMoveTime', 'ClickTime', 'randomtrain_startRating','Rating','reacttime','premovetime','ConfiScreenStartTime', 'ConfiMouseMoveTime', 'ConfiClickTime', 'randomtrain_confidencestartRating','confidenceRating','confi_reacttime','confi_premovetime','train_subADblock','train_subConditionblock','train_block'}; % 列标题
    rating_train_result = table(train_subinfo, train_picID, train_starttime, train_startA,train_startB,train_startC,train_startRating, train_Rating,train_RT_AC,train_RT_AB,train_confidencestartA,train_confidencestartB,train_confidencestartC,train_confidencestartRating,train_confidenceRating,train_confidenceRT_AC,train_confidenceRT_AB,train_subADblock,train_subConditionblock,train_block,'VariableNames', variableNames);
    rating_train_result = sortrows(rating_train_result, 'picID');

    subnum = subinfo{1};  % 从 cell 里取编号（可能是数字或字符串）
    baseDir = "E:\fMRI代码及其附属文件\实验数据\first_traindata\";

    % ⚙️ 转成字符串并拼接
    filename = fullfile(baseDir, sprintf('%s.xlsx', num2str(subnum)));

    % 检查路径是否正确
    disp(['文件将被保存到：', filename]);

    % 保存文件
    try
        writetable(rating_train_result, filename);
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
    WaitSecs(4);
    WaitSecs(slack);  % 控制刷新频率
end
end
