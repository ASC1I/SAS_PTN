
/*----------------*\
| Michael Lee      |
| 6/9/2017         |     
\*----------------*/

libname dataprep 'C:\SAS_Mastah\SAS Tableau';

*edit macrovariables here;
%let sheet = Marilyn_Jackson;   *<<<< select which sheet to import from excel file;
%let cy = 15;       *<<<< select starting year for a clinic;
%let endyear = 17;  *<<<< select ending year for a clinic;

/*---------------------------------------------------------*\
| run this macro to automatically import data into dataprep |
\*---------------------------------------------------------*/

%list_gen(clinic = &sheet);

*use this to check stats on the final report;
proc freq data= dataprep.final2;
	table outcome_measures * template_tab/ nocol norow nopercent;
run;

proc sql;
create table nigga as select distinct template_tab, outcome_measures from dataprep.final2;
quit;

data final;
set dataprep.final2;
format group $25.0;
if outcome_measures = "Diabetes: Urine Protein Screening" or 
outcome_measures = "Diabetes Mellitus (DM): Low Density Lipoprotein (LDL-C) Control" or
outcome_measures = "Diabetes HbA1c (poor control)" or
outcome_measures = "Blood pressure <140/90" 
then group = "diabetes";
else if outcome_measures = "Depression Screening and Follow-Up Plan" or
outcome_measures = "Avoidance of high risk medications" or
outcome_measures = "Breast cancer screening" or
outcome_measures = "Chlamydia Screening in Women" or
outcome_measures = "Cervical Cancer Screening" or
outcome_measures = "Colon cancer screening"
then group = "screening";
else if outcome_measures = "Appropriate Care for Children with Pharyngitis" or
outcome_measures = "Pediatrics: Childhood Immunizations" or
outcome_measures = "Well-child visits (age 3-6)" 
then group = "pediatrics";
else if outcome_measures = "Overall Doctor Rating" or
outcome_measures = "Access to care composite" or
outcome_measures = "Physician Communication Composite" or
outcome_measures = "Office Staff Quality Composite" or
outcome_measures = "Follow up on test results"
then group = "patient satisfaction";
run;


*imports specified sheet from excel file into sas;
*run this to import new sheets;
proc import out = Dataprep.&sheet. 
datafile = 'C:\SAS_Mastah\SAS Tableau\testbook'
dbms= xls replace; *<<<< the file must be in xls format, not xlsx!;
	sheet = "&sheet";
	namerow = 2;
	datarow = 3;
	getnames = no;
run;

proc import out = dataprep.final2
datafile = 'C:\SAS_Mastah\SAS Tableau\tableau20170613'
dbms = xlsx replace;
run;

/*creates a baseline from the original dataset for
future observations to be appeneded onto */

data dataprep.final (keep = qtr_start_dt clinic template_tab 
outcome_measures clinician practice num den rate);
	set dataprep.&sheet;
	qtr_start_dt = '01jan2015'd;
	format qtr_start_dt mmddyy10.;
	format clinic $30.;
	clinic = "&sheet";
	rename 
		bsl_clinician = clinician 
		bsl_practice = practice 
		bsl_num = num 
		bsl_den = den 
		bsl_rate_comp = rate;
	label
		bsl_clinician = clinician 
		bsl_practice = practice 
		bsl_num = num 
		bsl_den = den 
		bsl_rate_comp = rate;
run;

*this macro automatically determines quarter start date from a qid;
*is used to determine quarter start date in %slicer;
%macro date_gen(year =, qid = );
%do i = 15 %to &year;
	%if "&qid" = "CY&i.Q1" %then qtr_start_dt = mdy(01,01,&i);
	%else %if "&qid" = "CY&i.Q2" %then qtr_start_dt = mdy(04,01,&i);
	%else %if "&qid" = "CY&i.Q3" %then qtr_start_dt = mdy(07,01,&i);
	%else %if "&qid" = "CY&i.Q4" %then qtr_start_dt = mdy(10,01,&i);
%end;
%mend date_gen;

*this macro takes an id in the format of 'CY(year)(quarter)' and appends the corresponding data onto dataprep.final;
%macro slicer(cl= ,qy=, yr= );
data dataprep.temp (keep = clinic qtr_start_dt template_tab outcome_measures 
clinician practice num den rate);
set dataprep.&sheet;
format qtr_start_dt mmddyy10.;
clinic=symget('cl');
%date_gen(year = &yr, qid = &qy);
	rename  &qy._clinician = clinician
			&qy._practice = practice
			&qy._num = num
			&qy._den = den
			&qy._rate = rate;
run;

data dataprep.final;
	set dataprep.final dataprep.temp;
	if den = . then delete;  *<<<< any missing values are removed from the final report;
run;
%mend slicer;


*this macro generates the appropriate id number for %slicer to process;
%macro list_gen(clinic =);
%do i1 = &cy %to &endyear;
%do i2 = 1 %to 4;
%slicer(cl= &sheet ,qy= CY&i1.Q&i2, yr = &endyear);
%end;
%end;
%mend list_gen;

