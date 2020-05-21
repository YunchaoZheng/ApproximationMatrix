clear;
%-------------------DATA LOADING-------------------%
filename = '.\final.xls';
sheet = 1;
xlRange = 'C:F';
[searchPlus,text_total]=xlsread(filename,sheet,xlRange);
maxtmp = max(searchPlus)-0.000001;
mintmp = min(searchPlus)+0.000001;
searchPlus=(searchPlus-mintmp)/(maxtmp-mintmp);
[row,line]=size(text_total);
x_n = zeros(row-1,1);
for i = 2:1:row
    if(~isnan(str2double(cell2mat(text_total(i,1)))))
    x_n(i-1,1) = str2double(cell2mat(text_total(i,1)));
    end
end
maxtmp = max(x_n)-0.000001;
mintmp = min(x_n)+0.000001;
x_n=(x_n-mintmp)/(maxtmp-mintmp);
x_u = zeros(row-1,1);
for i = 2:1:row
    tmp = cell2mat(text_total(i,3));
    if(~isnan(str2double(tmp(2:length(tmp)-1))))
    x_u(i-1,1) = str2double(tmp(2:length(tmp)-1));
    end
end
maxtmp = max(x_u)-0.000001;
mintmp = min(x_u)+0.000001;
x_u=(x_u-mintmp)/(maxtmp-mintmp);
%{
for j = 1:1:line
    for i = 1:1:row
        container(i,j) = cell2mat(text_total(i,j));
    end
end
%}
%-------------------DATA LOADED-------------------%
%-------------------MODEL GENERATING-------------------%
%data not prepared *from PYTHON PART*
%initializing testing data
%{
not need any more
%}
%一元非线性回归
%这次没有时间做这个工作了，暂且输出数据用SPSS 22.0分析
title = ["searchPlus","x_n","x_u"];
xlswrite('.\out_spss.xls',[title;[searchPlus,x_n,x_u]]);
