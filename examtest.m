[~,~,dta]=xlsread('题库.xlsx');
% dta: qtpye,question,options,answer
dta=dta(2:end,2:end);

if ~isdir('pdf')
    mkdir('pdf');
end
savepath=[pwd,'\pdf\'];
filename=[savepath,'examtest.tex'];
fileID = fopen(filename,'w','n','UTF-8');


e='\\documentclass[printbox,marginline,noindent]{BHCexam}';
e=[e,char(10),'\\begin{document}',char(10)];
e=[e,'\\printmlol',char([10,10]),'\\maketitle',char(10)];
e=[e,'\\begin{questions}',char(10)];
fprintf(fileID,e);

% 填空题
tiankong_ind=find(cell2mat(dta(:,1))==1);
if ~isempty(tiankong_ind)
    nn=length(tiankong_ind);
    tiankong=['\tiankong',char(10)];
    for i=1:nn
        tiankong=[tiankong,char(10),'\question ',dta{tiankong_ind(i),2}];
        tiankong=[tiankong,'\stk{',dta{tiankong_ind(i),4},'}.',char(10)];
    end
end
tiankong = regexprep(tiankong,'\\','\\\');
fprintf(fileID,tiankong);

% 选择题
xuanze_ind=find(cell2mat(dta(:,1))==2);
if ~isempty(xuanze_ind)
    nn=length(xuanze_ind);
    xuanze=['\xuanze',char(10)];
    for i=1:nn
        xuanze=[xuanze,char(10),'\question ',dta{xuanze_ind(i),2}];
        xuanze=[xuanze,'\stk{',dta{xuanze_ind(i),4},'}.'];
        xuanze=[xuanze,char([10,10]),dta{xuanze_ind(i),3},char(10)];
    end
end
xuanze = regexprep(xuanze,'\\','\\\');
fprintf(fileID,xuanze);

% 简答题
question_ind=find(cell2mat(dta(:,1))==3);
if ~isempty(question_ind)
    nn=length(question_ind);
    question=['\jianda',char(10)];
    for i=1:nn
        question=[question,char(10),'\question ',dta{question_ind(i),2}];
        question=[question,'\stk{',dta{question_ind(i),4},'}.'];
        question=[question,char([10,10]),'\begin{solution}',...
            char(10),dta{question_ind(i),4},char(10),'\end{solution}',char(10)];
    end
end
question = regexprep(question,'\\','\\\');
fprintf(fileID,question);

% 结束
e=['\\end{questions}',char(10),'\\end{document}'];
fprintf(fileID,e);

fclose(fileID);

eval(['!xelatex ',filename]);
movefile('examtest.pdf',[savepath,'examtest.pdf']);
delete('examtest.aux','examtest.log');
winopen([savepath,'examtest.pdf']);


