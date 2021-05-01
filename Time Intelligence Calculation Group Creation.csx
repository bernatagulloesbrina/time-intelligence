// by Bernat Agulló
// www.esbrina-ba.com

//shout out to Johnny Winter for the base script and SQLBI for daxpatterns.com

//change the next 6 string variables for different naming conventions

//add the name of your calculation group here
string calcGroupName = "Time Intelligence";

//add the name for the column you want to appear in the calculation group
string columnName = "Time Calculation";


//add the name and date column name for fact table 
//over which you will create the measures to be used with the calculation group
string factTableName = "Sales";
string factTableDateColumnName = "Order Date";

//add the name for date table of the model
string dateTableName = "Date";
string dateTableDateColumnName = "Date";
string ShowValueForDatesMeasureName = "ShowValueForDates";
string dateWithSalesColumnName = "DateWithSales";



//
// ----- do not modify script below this line -----
//

string dateColumnWithTable = "'" + dateTableName + "'[" + dateTableDateColumnName + "]"; 
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
    "SELECTEDMEASURE()";


string PY = 
    "IF (" + 
    "    [" + ShowValueForDatesMeasureName + "], " + 
    "    CALCULATE ( " + 
    "        SELECTEDMEASURE(), " + 
    "        CALCULATETABLE ( " + 
    "            DATEADD ( " + dateColumnWithTable + " , -1, YEAR ), " + 
    "            " + dateWithSalesWithTable + " = TRUE " +  
    "        ) " + 
    "    ) " + 
    ") ";
    

string YOY = 
    "VAR ValueCurrentPeriod = SELECTEDMEASURE() " + 
    "VAR ValuePreviousPeriod = CALCULATE(SELECTEDMEASURE()," + calcGroupColumnWithTable + " = \"PY\" ) " +
    "VAR Result = " + 
    "IF ( " + 
    "    NOT ISBLANK ( ValueCurrentPeriod ) && NOT ISBLANK ( ValuePreviousPeriod ), " + 
    "     ValueCurrentPeriod - ValuePreviousPeriod" + 
    " ) " +  
    "RETURN " + 
    "   Result";

string YOYpct = 
    "DIVIDE ( "  + 
    "    CALCULATE(SELECTEDMEASURE()," + calcGroupColumnWithTable + " = \"YOY\" )," + 
    "    CALCULATE(SELECTEDMEASURE()," + calcGroupColumnWithTable + " = \"PY\" )" + 
    ")";
    
string YTD = 
    "IF (" +
    "    [" + ShowValueForDatesMeasureName + "]," + 
    "    CALCULATE (" +
    "        SELECTEDMEASURE()," + 
    "        DATESYTD (" +  dateColumnWithTable + " )" + 
    "   )" + 
    ")";
    
    
string PYTD = 
    "IF ( " + 
    "    [" + ShowValueForDatesMeasureName + "], " + 
    "   CALCULATE ( " + 
    "       CALCULATE(SELECTEDMEASURE()," + calcGroupColumnWithTable + " = \"YTD\" )," + 
    "    CALCULATETABLE ( " + 
    "        DATEADD ( " + dateColumnWithTable + ", -1, YEAR ), " + 
    "       " + dateWithSalesWithTable + " = TRUE " +  
    "       )" + 
    "   )" + 
    ")";
    

    
string YOYTD = 
    "VAR ValueCurrentPeriod = CALCULATE(SELECTEDMEASURE()," + calcGroupColumnWithTable + " = \"YTD\" )  " + 
    "VAR ValuePreviousPeriod = CALCULATE(SELECTEDMEASURE()," + calcGroupColumnWithTable + " = \"PYTD\" ) " +
    "VAR Result = " + 
    "IF ( " + 
    "    NOT ISBLANK ( ValueCurrentPeriod ) && NOT ISBLANK ( ValuePreviousPeriod ), " + 
    "     ValueCurrentPeriod - ValuePreviousPeriod" + 
    " ) " +  
    "RETURN " + 
    "   Result";


string YOYTDpct = 
    "DIVIDE ( "  + 
    "    CALCULATE(SELECTEDMEASURE()," + calcGroupColumnWithTable + " = \"YOYTD\" )," + 
    "    CALCULATE(SELECTEDMEASURE()," + calcGroupColumnWithTable + " = \"PYTD\" )" + 
    ")";
    

string[ , ] calcItems = 
    {
        {"CY",CY},
        {"PY",PY},
        {"YOY",YOY},
        {"YOY%",YOYpct},
        {"YTD",YTD},
        {"PYTD",PYTD},
        {"YOYTD",YOYTD},
        {"YOYTD%",YOYTDpct}
    };

    
int j = 0;

//create calculation items based on selected measures, including check to make sure calculation item doesnt exist
foreach(var cg in Model.CalculationGroups) {
    if (cg.Name == calcGroupName) {
        for (j = 0; j < calcItems.GetLength(0); j++) {
            
            string itemName = calcItems[j,0];
            string itemExpression = calcItems[j,1];
            
            if (!cg.CalculationItems.Contains(itemName)) {
                var nCalcItem = cg.AddCalculationItem(itemName, itemExpression);
                nCalcItem.FormatStringExpression = "SELECTEDMEASUREFORMATSTRING()";
                nCalcItem.FormatDax();
                nCalcItem.Ordinal = j; 
                
            };
        };
    };
};
