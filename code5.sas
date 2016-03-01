************************************************************
*the following code included data manipulation and analysis*
************************************************************

* data import and management;
PROC IMPORT OUT=cdata
DATAFILE="C:\Users\naz003\work\personal staff\New folder\765consulting\765consulting\process\originaldata\client.xls" 
DBMS=xls REPLACE;
GETNAMES=YES;
RUN;

data Cdata1;
set cdata;
if not(missing(dup)) ;
PREPDC_C=PRE_PDC;
POSPDC_C=POSTPDC;
if age>65 then Cat_age=1;
else Cat_age=0;
drop _6_MOS_AFT _6_MOS_PRE MTM_Date trans_PDC;
output;
run;
/*check the variables in the dataset*/
proc contents data=Cdata1;
run;
***************************************************************************;
/*Part 1.Comparison of the difference of PDC score */
/*two sample paired ttest for PREPDC_C and POSPDC_C*/
%let output=D:\school\Thesis\765consulting\process\output;
ods rtf file='&output\pttest.rtf' bodytitle  style=STATISTICAL;
proc ttest data=Cdata1; 
paired PREPDC_C*POSPDC_C; 
run; 
ods rtf close;
/*McNEMAR’S TEST for paired data*/
proc sql;
create table  PDCcom as
select input(VAR3,best12.) as pre_resp, input(VAR7,best12.) as pos_resp,Cat_age 
from Cdata1;
quit;
proc freq data=PDCcom;
table pre_resp*pos_resp/ nopercent nocum nocol norow out=freq;
run;
data freq1;
set freq;
drop PERCENT;
run;
ods rtf file='&output\Mcnemar.rtf' bodytitle  style=STATISTICAL;
proc freq data=freq1;
weight count; 
tables pre_resp*pos_resp/ agree ; 
exact mcnem ; 
run;
ods rtf close;
/*two sample ttest for different age subgroup*/
ods rtf file='&output\agetest.rtf' bodytitle  style=STATISTICAL;
proc ttest data=Cdata1;
class Cat_age;
var PDCdelta;
run;
ods rtf close;
******************************************************************************;
/*part 2 logistical regression */
data Cdata2;
set Cdata1;
IF VAR14=. THEN delete;
if DISTANCE=. then delete;
CMR=VAR15+1;
TIP=VAR16+1;
PREPDC=VAR3;
POSPDC=VAR7;
SIG=VAR14;
keep PREPDC POSPDC dup AGE Gender DISTANCE Distance_Category SIG Cat_age CMR TIP;
RUN;
data Cdata3; 
set Cdata2;
subject=_n_;
interven=0;
CMR_N=0;
TIP_N=0;
dup_N=0;
response=PREPDC;
output;
interven=1; 
dup_N=dup;
CMR_N=CMR;
TIP_N=TIP;
response =POSPDC;
output;
drop CMR TIP DUP PREPDC POSPDC;
run;
/*test the association between predictor variables and the PREPDC*/
%macro chi1(dataset,var1,var2,func1,func2,func3);
proc freq data=&dataset;
weight count;
table &var2*&var1/ &func1 nocol norow &func2 &func3;
run;
%mend chi;

data sig1;
input PREPDC $ sig count@@;
datalines;
G 1 28  L 1 0 
G 2 30  L 2 53
G 3 0   L 3 5
G 4 0  L 4 1
;
data gender1;
input PREPDC $ gender $ count@@;
datalines;
G F 39  L F 33 
G M 19  L M 26
;

data age1;
input PREPDC $ category_age count@@;
datalines;
G 0 1  L 0 9 
G 1 57  L 1 50
;
data distance1;
input PREPDC $ category_distance count@@;
datalines;
G 1 36  L 1 28 
G 2 15  L 2 22
G 3 7   L 3 9
;
ods rtf file='&output\freq_table.rtf' bodytitle  style=STATISTICAL;
%chi1(SIG1,sig,PREPDC,chisq,trend)
%chi1(GENDER1,Gender,prePDC,chisq)
%chi1(age1,category_age,prePDC,chisq)
%chi1(distance1,Category_Distance,prePDC,chisq,cmh)
ods rtf close;


/*logistic regression full model between response variable PREPDC and SIG,Cate_age*/
ods rtf file='&output\fullmodle_prepdc.rtf' bodytitle  style=STATISTICAL;
proc logistic data=Cdata2 descending;
class Cat_age(ref='0') SIG(REF='1') /param=ref;
model PREPDC= Cat_age SIG Cat_age*SIG;
EXACT intercept SIG Cat_age Cat_age*SIG/estimate=both;
run;
ods rtf close;
/*logistic regression reduce model without interaction SIG*Cat_age*/
ods rtf file='&output\mainmodle_prepdc.rtf' bodytitle  style=STATISTICAL;
proc logistic data=Cdata2 descending plots(only)=oddsratio(logbase=2);
class Cat_age(ref='0') SIG(REF='1') /param=ref;
model PREPDC= Cat_age SIG;
EXACT intercept SIG Cat_age/estimate=both; 
run;
ods rtf close;
/*test association between predictor variables and POSPDC*/
data sig2;
input POSPDC $ sig count@@;
datalines;
G 1 27  L 1 1 
G 2 37  L 2 46
G 3 3   L 3 2
G 4 1  L 4 0
;
data age2;
input POSPDC $ category_age count@@;
datalines;
G 0 3  L 0 7 
G 1 65  L 1 42
;
data gender2;
input PosPDC $ gender $ count@@;
datalines;
G F 40  L F 32 
G M 28  L M 17
;

data distance2;
input POSPDC $ Category_Distance count@@;
datalines;
G 1 43  L 1 21 
G 2 18  L 2 19
G 3 7   L 3 9
;

data dup2;
input pospdc dup count@@;
datalines;
0 1 54 1 1 28 
0 2 9  1 2 10
0 3 5  1 3 11
;

data CMR;
input pospdc $ cmr count@@;
datalines;
g 0 65 l 0 47 
g 1 3  l 1 2
;

data tip;
input pospdc $ TIP count@@;
datalines;
G 0 65 L 0 47 
G 1 3  L 1 2
;

ods rtf file='&output\sig_pospdc.rtf' bodytitle  style=STATISTICAL;

%chi1(SIG2,sig,POSPDC,chisq,trend,cmh)
%chi1(AGE2,category_age,POSPDC,chisq)
%chi1(gender2,gender,POSPDC,chisq)
%chi1(distance2,Category_Distance,POSPDC,chisq)
%chi1(dup2,dup,POSPDC,chisq,cmh,trend)
%chi1(cmr,CMR,POSPDC,chisq,cmh)
%chi1(tip,tip,POSPDC,chisq,cmh)
ods rtf close;


/*full logistic regression for response variable POSPDC and predictor variables dup, cate-distance and sig*/
ods rtf file='&output\fullmodel_pospdc.rtf' bodytitle  style=STATISTICAL;
proc logistic data=Cdata2 descending;
class Cat_age(ref='0') dup(ref='1') tip(ref='1') Distance_Category(ref='1') SIG(REF='1') /param=ref;
model PosPDC= dup|Distance_Category|SIG@2;
EXACT intercept   dup  Distance_Category SIG; 
run;
ods rtf close;
/*the main effect model for posPDC*/
ods rtf file='&output\mainmodel_pospdc.rtf' bodytitle  style=STATISTICAL;
proc logistic data=Cdata2 descending;
class Cat_age(ref='0') dup(ref='3') tip(ref='1') Distance_Category(ref='1') SIG(REF='1') /param=ref;
model PosPDC=dup Cat_age  tip Distance_Category SIG/selection=forward include=1 ;
run;
ods rtf close;
/*the full model of logostic conditional regression*/
ods rtf file='&output\conditional_full.rtf' bodytitle  style=STATISTICAL;
proc logistic data=Cdata3 descending;   
   class interven(ref='0') CMR_N(ref='0') TIP_N(ref='0') dup_N(ref='0')/ param=ref; 
   strata subject;
   model response =interven|dup_N|CMR_N|TIP_N@2/selection=forward include=1 details;
  
run;
ods rtf close;
/*the reduce model of logostic conditional regression*/
proc logistic data=Cdata3 descending;  
   class interven(ref='0') TIP_N(ref='0')/ param=ref; 
   strata subject;
   model response=interven TIP_N;
run;
/*the final conditional model include the intervention*/
ods rtf file='&output\finalmodel.rtf' bodytitle  style=STATISTICAL;
ods graphics on;
proc logistic data=Cdata3 descending plots(only)=oddsratio(logbase=2);  
   class interven(ref='0')  / param=ref; 
   strata subject;
   model response =interven ;
   oddsratios interven/ DIFF=REF;
run;
ods graphics off;
ods rtf close;
*************************************************************************************************;
/*Part 3 linear regression*/
data Cdata4;
set Cdata;
ratiopdc=POSTPDC/PRE_PDC;
if ratiopdc=<0 then delete;
if VAR14=. then delete;
if DISTANCE=. then delete;
logpdc=log(ratiopdc);
sig=var14;
CMR=var15;
TIP=var16;
logage=log(age);
logdistance=log(distance);
if gender='F' THEN ngender=0;
else ngender=1;
keep  dup AGE ngender distance sig CMR TIP ratiopdc logpdc  PDCdelta logage logdistance;
run;
/*plot the PDCdelta distribution*/


%macro univar(dataset,var);
proc univariate data=&dataset;
   var &var ;
   histogram / normal(color=red)midpoints=-0.7 to 0.7 by 0.05;
    inset n mean(5.3) std='Std Dev'(5.3) skewness(5.3) kurtosis(5.3)
       / pos = ne  header = 'Summary Statistics';
qqplot;
run;
%mend univar;
ods rtf file='&output\univariate.rtf' bodytitle  style=STATISTICAL;
%univar(cdata4,PDCdelta)
%univar(cdata4,ratiopdc)
ods rtf close;

/*transform the ratiopdc*/
ods rtf file='&output\boxcox.rtf' bodytitle  style=STATISTICAL;
proc transreg data=cdata4 ss2 details;
title2 'Defaults';
model boxcox(ratioPDC) = identity(tip sig CMR dup AGE  DISTANCE ngender);
run;
ods rtf close;

proc IML;
use Cdata4;
read all var "ratioPDC" into y;
lny=log(y);
n=nrow(y);
one=j(n,1,1);
avglog=lny`*one/n;
print avglog;
geomean=exp(avglog);
print geomean;
geomeany=geomean#one;
create gmean var {geomeany};
append from geomeany;
close gmean;
quit;
run;

data cdata5;
merge Cdata4 gmean;
transpdc=((ratioPDC**(0.75))-1)/(0.75*(geomeany**(0.75-1)));
n=_n_;
run;
/*plot the transPDC distribution*/
ods rtf file='&output\transratio_pdf.rtf' bodytitle  style=STATISTICAL;
%univar(cdata5,transpdc)
ods rtf close;

/*Data Diagnostics*/
/*test correlation between predictor variables and transPDC*/
ods rtf file='&output\correlation.rtf' bodytitle  style=STATISTICAL;
 ODS GRAPHICS ON;
PROC CORR DATA=Cdata5 noprob plots( MAXPOINTS=NONE)=matrix(NVAR=ALL) ;
 VAR transpdc sig CMR TIP age Distance dup ngender;
 RUN;
ods rtf close;
/*produce the studentized residuals,leverage,cook's D,deffit,debeta */
 PROC REG DATA=cdata5;
 MODEL  logpdc=tip sig CMR dup age distance ngender ;
 OUTPUT OUT=OUT rstudent=r h=lev cookd=cd dffits=dffit;
 RUN;
 quit;
/*plot the correlation between presictor variables and residuals*/
goptions reset=all;
axis1 label=(r=0 a=90);
symbol1 pointlabel = ("#n") font=simplex value=none;

%macro ggplot(dataset,var);
proc gplot DATA=&dataset;
plot r*&var;
RUN;
quit;
%mend;
ods rtf file='&output\X_residual.rtf' bodytitle  style=STATISTICAL;
ODS GRAPHICS ON;
%ggplot(out,dup)
%ggplot(out,age)
%ggplot(out,nGender)
%ggplot(out,distance)
%ggplot(out,sig)
%ggplot(out,CMR)
%ggplot(out,TIP)
ods rtf close;
/*find out the abnormal studentized residual*/
%macro printer(dataset,var1,var2);
proc print data=&dataset;
  var &var1 &var2  n;
run;
%mend printer;

ods rtf file='&output\SR-abnormal.rtf' bodytitle  style=STATISTICAL;
%printer(out,r,transpdc)
%printer(out,lev,transpdc)
%printer(out,cd,transpdc)
%printer(out,dffit,transpdc)
ods rtf close;


/*find out for abnormal DFBeta observations*/
proc reg data=Cdata5;
  model transpdc=tip sig CMR dup age distance ngender/ influence;
  ods output OutputStatistics=cdatadfbetas;
run;
quit;

data cdatadfbetas1;
set cdatadfbetas;
n=_n_;
run;

%macro printer1(dataset,var1);
proc print data=&dataset;
where abs(DFB_TIP)> 2/sqrt(117);
var &var1 n;
run;
%mend printer1;

ods rtf file='&output\printe_out.rtf' bodytitle  style=STATISTICAL;
%printer1(cdatadfbetas1,DFB_TIP)
%printer1(cdatadfbetas1,DFB_sig)
%printer1(cdatadfbetas1,DFB_CMR)
%printer1(cdatadfbetas1,DFB_DUP)
%printer1(cdatadfbetas1,dfb_age)
%printer1(cdatadfbetas1,DFB_distance)
%printer1(cdatadfbetas1,dfb_ngender)
ods rtf close;
/*plot the influential and outlier observations using the criterion calculated above*/
ods rtf file='&output\summary_diagosis_plot.rtf' bodytitle  style=STATISTICAL;
 PROC REG DATA=Cdata5 plots=(diagnostics(stats=none) RStudentByLeverage(label) 
              CooksD(label) Residuals(smooth)
              DFFITS(label) DFBETAS ObservedByPredicted(label));;
 MODEL transpdc=tip sig CMR dup AGE Distance ngender/ vif r INFLUENCE;
 RUN;
 quit;
 ods rtf close;
/*delete outlier n=30*/
data cdata6;
set cdata5;
if n=30 then delete;
run;
/*recheck the outlier*/
ods rtf file='&output\recheckoutlier.rtf' bodytitle  style=STATISTICAL;
 PROC REG DATA=Cdata6 plots=(diagnostics(stats=none) RStudentByLeverage(label) 
              CooksD(label) Residuals(smooth)
              DFFITS(label) DFBETAS ObservedByPredicted(label));;
 MODEL transpdc=tip sig CMR dup AGE distance ngender/ vif r INFLUENCE;
 RUN;
 quit;
 ods rtf close;
/*model selection*/
ods rtf file='&output\selection.rtf' bodytitle  style=STATISTICAL;
proc reg data=Cdata6;
 model transpdc=tip sig CMR dup AGE distance ngender/selection=forward ;
run;
quit;
ods rtf close;
/*Tests Normality for Residuals*/
proc reg data=Cdata6;
 model transpdc=tip CMR dup;
output out=elem1res(keep=transpdc tip CMR dup r fv) residual=r predicted=fv;
run;
quit;

ods rtf file='&output\residual_qq.rtf' bodytitle  style=STATISTICAL;
goptions reset=all;
%univar(elem1res,r)
ods rtf close;
/*test for homogeneity*/
ods rtf file='&output\homogeneity.rtf' bodytitle  style=STATISTICAL;
proc reg data=Cdata6;
 model transpdc=tip CMR dup;
 plot r.*p.;
 plot cookd.*obs.;
run;
quit;
ods rtf close;
/*test for the collinearity,vif and tol*/
ods rtf file='&output\collinearity.rtf' bodytitle  style=STATISTICAL;
proc reg data=Cdata6;
 model transpdc=tip CMR dup/vif tol collinoint;
run;
quit;
ods rtf close;
/*test the linearty of seleced predictory variables and response variable transPDC*/
ods rtf file='&output\dup_Y_linearty.rtf' bodytitle  style=STATISTICAL;
proc reg data=Cdata6;
 model transpdc=dup ;
 plot rstudent.*p. / noline;
 plot transpdc*dup;
run;
quit;
ODS RTF CLOSE;

ods rtf file='&output\CMR_Y_linearty.rtf' bodytitle  style=STATISTICAL;
proc reg data=Cdata6;
 model transpdc=CMR ;
 plot rstudent.*p. / noline;
 plot transpdc*CMR;
run;
quit;
ODS RTF CLOSE;

ods rtf file='&output\tip_Y_linearty.rtf' bodytitle  style=STATISTICAL;
proc reg data=Cdata6;
 model transpdc=tip;
 plot rstudent.*p. / noline;
 plot transpdc*tip;
run;
quit;
ods rtf close;
/*test the indepently*/
ods rtf file='&output\DW_independence.rtf' bodytitle  style=STATISTICAL;
proc reg data=Cdata6;
 model transpdc=tip cmr dup / dw;
 output out=res3  residual=r ;
run;
quit;
ODS rtf close;
/*estimate the parameters in the model*/
data cdata7;
  set cdata6;
    if dup~=.  then dup1=0;
    if dup~=.  then dup2=0;
    if dup~=.  then dup3=0;
    if dup = 1 then dup1=1;
    if dup = 2 then dup2=1;
    if dup = 3 then dup3=1;
run;
/*estimate the parameter*/
proc reg data = Cdata7;
      model transpdc=tip cmr  dup2 dup3 /acov;
      ods output  ACovEst = estcov;
      ods output ParameterEstimates=pest;
run;
quit;
/*because of the heterogeneity exsiting among residuals, using onther method to calculate the adjust stardard error for parameters*/
data temp_dm;
  set estcov;
  drop model dependent;
  array a(5) intercept tip cmr dup1 dup2;
  array b(5) std1-std5;
  b(_n_) = sqrt((116/111)*a(_n_));
  std = max(of std1-std4);
  keep variable std;
run;
/*produce the new standard error for each parameter estimate*/
ods rtf file='&output\stardarderror.rtf' bodytitle  style=STATISTICAL;
proc sql;
  select pest.variable, estimate, stderr, tvalue, probt, std as robust_stderr, 
         estimate/robust_stderr as tvalue_rb, 
         (1 - probt(abs(estimate/robust_stderr), 115))*2 as probt_rb 
  from pest, temp_dm
  where pest.variable=temp_dm.variable;
quit;
ods rtf close;


