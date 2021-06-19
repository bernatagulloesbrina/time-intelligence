// '2021-05-01 / B.Agullo / 
// '2021-05-17 / B.Agullo / added affected measure table
// '2021-06-19 / B.Agullo / data label measures
// by Bernat Agulló
// www.esbrina-ba.com

//shout out to Johnny Winter for the base script and SQLBI for daxpatterns.com

//select the measures that you want to be affected by the calculation group
//before running the script. 
//measure names can also be included in the following array (no need to select them) 
string[] preSelectedMeasures = {}; //include measure names in double quotes, like: {"Profit","Total Cost"};

//AT LEAST ONE MEASURE HAS TO BE AFFECTED, 
//either by selecting it or typing its name in the preSelectedMeasures Variable

//change the next string variables to fit your model

//add the name of your calculation group here
string calcGroupName = "Time Intelligence";

//add the name for the column you want to appear in the calculation group
string columnName = "Time Calculation";

//add the name and date column name for fact table 
string factTableName = "Sales";
string factTableDateColumnName = "Order Date";

//add the name for date table of the model
string dateTableName = "Date";
string dateTableDateColumnName = "Date";
string dateTableYearColumnName = "CalendarYear"; 

//add the measure and calculated column names you want or leave them as they are
string ShowValueForDatesMeasureName = "ShowValueForDates";
string dateWithSalesColumnName = "DateWithSales";

string affectedMeasuresTableName = "Time Intelligence Affected Measures"; 
string affectedMeasuresColumnName = "Measure"; 


string labelAsValueMeasureName = "Label as Value Measure"; 
string labelAsFormatStringMeasureName = "Label as format string"; 



//
// ----- do not modify script below this line -----
//


string calcItemProtection = "<CODE>"; //default value if user has selected no measures
string calcItemFormatProtection = "<CODE>"; //default value if user has selected no measures

string affectedMeasures = "{";

int i = 0; 

for (i=0;i<preSelectedMeasures.GetLength(0);i++){
  
    if(affectedMeasures == "{") {
    affectedMeasures = affectedMeasures + "\"" + preSelectedMeasures[i] + "\"";
    }else{
        affectedMeasures = affectedMeasures + ",\"" + preSelectedMeasures[i] + "\"" ;
    }; 
    
};


if (Selected.Measures.Count != 0) {
    
    foreach(var m in Selected.Measures) {
        if(affectedMeasures == "{") {
        affectedMeasures = affectedMeasures + "\"" + m.Name + "\"";
        }else{
            affectedMeasures = affectedMeasures + ",\"" + m.Name + "\"" ;
        };
    };  
};

//check that by either method at least one measure is affected
if(affectedMeasures == "{") { 
    Error("No measures affected by calc group"); 
    return; 
};

//if there where selected or preselected measures, prepare protection code for expresion and formatstring
if(affectedMeasures != "{") { 
    
    affectedMeasures = affectedMeasures + "}";
    
    string affectedMeasureTableExpression = 
        "SELECTCOLUMNS(" + affectedMeasures + ",\"" + affectedMeasuresColumnName + "\",[Value])";

    var affectedMeasureTable = 
        Model.AddCalculatedTable(affectedMeasuresTableName,affectedMeasureTableExpression);
    
    affectedMeasureTable.FormatDax(); 
    affectedMeasureTable.Description = 
        "Measures affected by " + calcGroupName + " calculation group." ;
    
    affectedMeasureTable.IsHidden = true;     
    
    string affectedMeasuresValues = "VALUES('" + affectedMeasuresTableName + "'[" + affectedMeasuresColumnName + "])";
    
    calcItemProtection = 
        "SWITCH(" + 
        "   TRUE()," + 
        "   SELECTEDMEASURENAME() IN " + affectedMeasuresValues + "," + 
        "   <CODE> ," + 
        "   ISSELECTEDMEASURE([" + labelAsValueMeasureName + "])," + 
        "   <LABELCODE> ," + 
        "   SELECTEDMEASURE() " + 
        ")";
        
        
    calcItemFormatProtection = 
        "SWITCH(" + 
        "   TRUE() ," + 
        "   SELECTEDMEASURENAME() IN " + affectedMeasuresValues + "," + 
        "   <CODE> ," + 
        "   ISSELECTEDMEASURE([" + labelAsFormatStringMeasureName + "])," + 
        "   <LABELCODEFORMATSTRING> ," +
        "   SELECTEDMEASUREFORMATSTRING() " + 
        ")";
};
    
string dateColumnWithTable = "'" + dateTableName + "'[" + dateTableDateColumnName + "]"; 
string yearColumnWithTable = "'" + dateTableName + "'[" + dateTableYearColumnName + "]"; 
string factDateColumnWithTable = "'" + factTableName + "'[" + factTableDateColumnName + "]";
string dateWithSalesWithTable = "'" + dateTableName + "'[" + dateWithSalesColumnName + "]";
string calcGroupColumnWithTable = "'" + calcGroupName + "'[" + columnName + "]";

//check to see if a table with this name already exists
//if it doesnt exist, create a calculation group with this name
if (!Model.Tables.Contains(calcGroupName)) {
  var cg = Model.AddCalculationGroup(calcGroupName);
  cg.Description = "Calculation group for time intelligence. Availability of data is taken from " + factTableName + ".";
};

//set variable for the calc group
Table calcGroup = Model.Tables[calcGroupName];

//if table already exists, make sure it is a Calculation Group type
if (calcGroup.SourceType.ToString() != "CalculationGroup") {
  Error("Table exists in Model but is not a Calculation Group. Rename the existing table or choose an alternative name for your Calculation Group.");
  return;
};

//adds the two measures that will be used for label as value, label as format string 
var labelAsValueMeasure = calcGroup.AddMeasure(labelAsValueMeasureName,"0");
labelAsValueMeasure.Description = "Use this measure to show the year evaluated in tables"; 

var labelAsFormatStringMeasure = calcGroup.AddMeasure(labelAsFormatStringMeasureName,"0");
labelAsFormatStringMeasure.Description = "Use this measure to show the year evaluated in charts"; 

//by default the calc group has a column called Name. If this column is still called Name change this in line with specfied variable
if (calcGroup.Columns.Contains("Name")) {
  calcGroup.Columns["Name"].Name = columnName;
};

calcGroup.Columns[columnName].Description = "Select value(s) from this column to apply time intelligence calculations.";

//set variable for the date table 
Table dateTable = Model.Tables[dateTableName];


string DateWithSalesCalculatedColumnExpression = 
    dateColumnWithTable + " <= MAX ( " + factDateColumnWithTable + ")";

dateTable.AddCalculatedColumn(dateWithSalesColumnName,DateWithSalesCalculatedColumnExpression);


string ShowValueForDatesMeasureExpression = 
    "VAR LastDateWithData = " + 
    "    CALCULATE ( " + 
    "        MAX (  " + factDateColumnWithTable + " ), " + 
    "        REMOVEFILTERS () " +
    "    )" +
    "VAR FirstDateVisible = " +
    "    MIN ( " + dateColumnWithTable + " ) " + 
    "VAR Result = " +  
    "    FirstDateVisible <= LastDateWithData " +
    "RETURN " + 
    "    Result ";

var ShowValueForDatesMeasure = dateTable.AddMeasure(ShowValueForDatesMeasureName,ShowValueForDatesMeasureExpression); 

ShowValueForDatesMeasure.FormatDax();


string CY = 
    "/*CY*/ " + 
    "SELECTEDMEASURE()";

string CYlabel = 
    "SELECTEDVALUE(" + yearColumnWithTable + ")";


string PY = 
    "/*PY*/ " +
    "IF (" + 
    "    [" + ShowValueForDatesMeasureName + "], " + 
    "    CALCULATE ( " + 
    "        "+ CY + ", " + 
    "        CALCULATETABLE ( " + 
    "            DATEADD ( " + dateColumnWithTable + " , -1, YEAR ), " + 
    "            " + dateWithSalesWithTable + " = TRUE " +  
    "        ) " + 
    "    ) " + 
    ") ";
    

string PYlabel = 
    "/*PY*/ " +
    "IF (" + 
    "    [" + ShowValueForDatesMeasureName + "], " + 
    "    CALCULATE ( " + 
    "        "+ CYlabel + ", " + 
    "        CALCULATETABLE ( " + 
    "            DATEADD ( " + dateColumnWithTable + " , -1, YEAR ), " + 
    "            " + dateWithSalesWithTable + " = TRUE " +  
    "        ) " + 
    "    ) " + 
    ") ";   


string YOY = 
    "/*YOY*/ " + 
    "VAR ValueCurrentPeriod = " + CY + " " + 
    "VAR ValuePreviousPeriod = " + PY + " " +
    "VAR Result = " + 
    "IF ( " + 
    "    NOT ISBLANK ( ValueCurrentPeriod ) && NOT ISBLANK ( ValuePreviousPeriod ), " + 
    "     ValueCurrentPeriod - ValuePreviousPeriod" + 
    " ) " +  
    "RETURN " + 
    "   Result ";

string YOYlabel = 
    "/*YOY*/ " + 
    "VAR ValueCurrentPeriod = " + CYlabel + " " + 
    "VAR ValuePreviousPeriod = " + PYlabel + " " +
    "VAR Result = " + 
    "IF ( " + 
    "    NOT ISBLANK ( ValueCurrentPeriod ) && NOT ISBLANK ( ValuePreviousPeriod ), " + 
    "     ValueCurrentPeriod & \" vs \" & ValuePreviousPeriod" + 
    " ) " +  
    "RETURN " + 
    "   Result ";

string YOYpct = 
    "/*YOY%*/ " +
   "VAR ValueCurrentPeriod = " + CY + " " + 
    "VAR ValuePreviousPeriod = " + PY + " " + 
    "VAR CurrentMinusPreviousPeriod = " +
    "IF ( " + 
    "    NOT ISBLANK ( ValueCurrentPeriod ) && NOT ISBLANK ( ValuePreviousPeriod ), " + 
    "     ValueCurrentPeriod - ValuePreviousPeriod" + 
    " ) " +  
    "VAR Result = " + 
    "DIVIDE ( "  + 
    "    CurrentMinusPreviousPeriod," + 
    "    ValuePreviousPeriod" + 
    ") " + 
    "RETURN " + 
    "  Result";

string YOYpctLabel = 
    "/*YOY%*/ " +
   "VAR ValueCurrentPeriod = " + CYlabel + " " + 
    "VAR ValuePreviousPeriod = " + PYlabel + " " + 
    "VAR Result = " +
    "IF ( " + 
    "    NOT ISBLANK ( ValueCurrentPeriod ) && NOT ISBLANK ( ValuePreviousPeriod ), " + 
    "     ValueCurrentPeriod & \" vs \" & ValuePreviousPeriod & \" (%)\"" + 
    " ) " +  
    "RETURN " + 
    "  Result";
    
string YTD = 
    "/*YTD*/" + 
    "IF (" +
    "    [" + ShowValueForDatesMeasureName + "]," + 
    "    CALCULATE (" +
    "        " + CY+ "," + 
    "        DATESYTD (" +  dateColumnWithTable + " )" + 
    "   )" + 
    ") ";
    

string YTDlabel = CYlabel + "& \" YTD\""; 


string PYTD = 
    "/*PYTD*/" + 
    "IF ( " + 
    "    [" + ShowValueForDatesMeasureName + "], " + 
    "   CALCULATE ( " + 
    "       " + YTD + "," + 
    "    CALCULATETABLE ( " + 
    "        DATEADD ( " + dateColumnWithTable + ", -1, YEAR ), " + 
    "       " + dateWithSalesWithTable + " = TRUE " +  
    "       )" + 
    "   )" + 
    ") ";
    
string PYTDlabel = PYlabel + "& \" YTD\""; 

    
string YOYTD = 
    "/*YOYTD*/" + 
    "VAR ValueCurrentPeriod = " + YTD + " " + 
    "VAR ValuePreviousPeriod = " + PYTD + " " + 
    "VAR Result = " + 
    "IF ( " + 
    "    NOT ISBLANK ( ValueCurrentPeriod ) && NOT ISBLANK ( ValuePreviousPeriod ), " + 
    "     ValueCurrentPeriod - ValuePreviousPeriod" + 
    " ) " +  
    "RETURN " + 
    "   Result ";


string YOYTDlabel = 
    "/*YOYTD*/" + 
    "VAR ValueCurrentPeriod = " + YTDlabel + " " + 
    "VAR ValuePreviousPeriod = " + PYTDlabel + " " + 
    "VAR Result = " + 
    "IF ( " + 
    "    NOT ISBLANK ( ValueCurrentPeriod ) && NOT ISBLANK ( ValuePreviousPeriod ), " + 
    "     ValueCurrentPeriod & \" vs \" & ValuePreviousPeriod" + 
    " ) " +  
    "RETURN " + 
    "   Result ";



string YOYTDpct = 
    "/*YOYTD%*/" + 
    "VAR ValueCurrentPeriod = " + YTD + " " + 
    "VAR ValuePreviousPeriod = " + PYTD + " " + 
    "VAR CurrentMinusPreviousPeriod = " +
    "IF ( " + 
    "    NOT ISBLANK ( ValueCurrentPeriod ) && NOT ISBLANK ( ValuePreviousPeriod ), " + 
    "     ValueCurrentPeriod - ValuePreviousPeriod" + 
    " ) " +  
    "VAR Result = " + 
    "DIVIDE ( "  + 
    "    CurrentMinusPreviousPeriod," + 
    "    ValuePreviousPeriod" + 
    ") " + 
    "RETURN " + 
    "  Result";


string YOYTDpctLabel = 
    "/*YOY%*/ " +
   "VAR ValueCurrentPeriod = " + YTDlabel + " " + 
    "VAR ValuePreviousPeriod = " + PYTDlabel + " " + 
    "VAR Result = " +
    "IF ( " + 
    "    NOT ISBLANK ( ValueCurrentPeriod ) && NOT ISBLANK ( ValuePreviousPeriod ), " + 
    "     ValueCurrentPeriod & \" vs \" & ValuePreviousPeriod & \" (%)\"" + 
    " ) " +  
    "RETURN " + 
    "  Result";
    

string defFormatString = "SELECTEDMEASUREFORMATSTRING()";
string pctFormatString = "\"#,##0.#%\"";


//the order in the array also determines the ordinal position of the item    
string[ , ] calcItems = 
    {
        {"CY",      CY,         defFormatString,    "Current year",             CYlabel},
        {"PY",      PY,         defFormatString,    "Previous year",            PYlabel},
        {"YOY",     YOY,        defFormatString,    "Year-over-year",           YOYlabel},
        {"YOY%",    YOYpct,     pctFormatString,    "Year-over-year%",          YOYpctLabel},
        {"YTD",     YTD,        defFormatString,    "Year-to-date",             YTDlabel},
        {"PYTD",    PYTD,       defFormatString,    "Previous year-to-date",    PYTDlabel},
        {"YOYTD",   YOYTD,      defFormatString,    "Year-over-year-to-date",   YOYTDlabel},
        {"YOYTD%",  YOYTDpct,   pctFormatString,    "Year-over-year-to-date%",  YOYTDpctLabel},
    };

    
int j = 0;


//create calculation items for each calculation with formatstring and description
foreach(var cg in Model.CalculationGroups) {
    if (cg.Name == calcGroupName) {
        for (j = 0; j < calcItems.GetLength(0); j++) {
            
            string itemName = calcItems[j,0];
            
            string itemExpression = calcItemProtection.Replace("<CODE>",calcItems[j,1]);
            itemExpression = itemExpression.Replace("<LABELCODE>",calcItems[j,4]); 
            
            string itemFormatExpression = calcItemFormatProtection.Replace("<CODE>",calcItems[j,2]);
            itemFormatExpression = itemFormatExpression.Replace("<LABELCODEFORMATSTRING>","\"\"\"\" & " + calcItems[j,4] + " & \"\"\"\"");
            
            //if(calcItems[j,2] != defFormatString) {
            //    itemFormatExpression = calcItemFormatProtection.Replace("<CODE>",calcItems[j,2]);
            //};

            string itemDescription = calcItems[j,3];
            
            if (!cg.CalculationItems.Contains(itemName)) {
                var nCalcItem = cg.AddCalculationItem(itemName, itemExpression);
                nCalcItem.FormatStringExpression = itemFormatExpression;
                nCalcItem.FormatDax();
                nCalcItem.Ordinal = j; 
                nCalcItem.Description = itemDescription;
                
            };
        };


    };
};
