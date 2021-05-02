// by Bernat Agulló
// www.esbrina-ba.com

//shout out to Johnny Winter for the base script and SQLBI for daxpatterns.com

//change the next 6 string variables for different naming conventions

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

//add the measure and calculated column names you want or leave them as they are
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
    "/*CY*/ " + 
    "SELECTEDMEASURE()";


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
    
string YTD = 
    "/*YTD*/" + 
    "IF (" +
    "    [" + ShowValueForDatesMeasureName + "]," + 
    "    CALCULATE (" +
    "        " + CY+ "," + 
    "        DATESYTD (" +  dateColumnWithTable + " )" + 
    "   )" + 
    ") ";
    
    
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
    

    
string YOYTD = 
    "/*YOYTD*/" + 
    "VAR ValueCurrentPeriod = " + YTD + 
    "VAR ValuePreviousPeriod = " + PYTD +
    "VAR Result = " + 
    "IF ( " + 
    "    NOT ISBLANK ( ValueCurrentPeriod ) && NOT ISBLANK ( ValuePreviousPeriod ), " + 
    "     ValueCurrentPeriod - ValuePreviousPeriod" + 
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
    

string defFormatString = "SELECTEDMEASUREFORMATSTRING()";
string pctFormatString = "\"#,##0.#%\"";


//the order in the array also determines the ordinal position of the item    
string[ , ] calcItems = 
    {
        {"CY",      CY,         defFormatString,    "Current year"},
        {"PY",      PY,         defFormatString,    "Previous year"},
        {"YOY",     YOY,        defFormatString,    "Year-over-year" },
        {"YOY%",    YOYpct,     pctFormatString,    "Year-over-year%"},
        {"YTD",     YTD,        defFormatString,    "Year-to-date"},
        {"PYTD",    PYTD,       defFormatString,    "Previous year-to-date"},
        {"YOYTD",   YOYTD,      defFormatString,    "Year-over-year-to-date"},
        {"YOYTD%",  YOYTDpct,   pctFormatString,    "Year-over-year-to-date%"},
    };

    
int j = 0;

//create calculation items for each calculation with formatstring and description
foreach(var cg in Model.CalculationGroups) {
    if (cg.Name == calcGroupName) {
        for (j = 0; j < calcItems.GetLength(0); j++) {
            
            string itemName = calcItems[j,0];
            string itemExpression = calcItems[j,1];
            string itemFormatExpression = calcItems[j,2];
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
