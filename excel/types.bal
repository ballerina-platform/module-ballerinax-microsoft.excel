// Copyright (c) 2021, WSO2 Inc. (http://www.wso2.org) All Rights Reserved.
//
// WSO2 Inc. licenses this file to you under the Apache License,
// Version 2.0 (the "License"); you may not use this file except
// in compliance with the License.
// You may obtain a copy of the License at
//
// http://www.apache.org/licenses/LICENSE-2.0
//
// Unless required by applicable law or agreed to in writing,
// software distributed under the License is distributed on an
// "AS IS" BASIS, WITHOUT WARRANTIES OR CONDITIONS OF ANY
// KIND, either express or implied.  See the License for the
// specific language governing permissions and limitations
// under the License.

# Represents worksheet properties
#
# + id - Returns a value that uniquely identifies the worksheet in a given workbook. The value of the identifier remains
# the same even when the worksheet is renamed or moved
# + position - The zero-based position of the worksheet within the workbook
# + name - Worksheet name
# + visibility - The visibility of the worksheet
@display {label: "Worksheet"}
public type Worksheet record {
    string & readonly id?;
    @display {label: "Position"}
    int position?;
    @display {label: "Worksheet Name"}
    string name?;
    Visibility visibility?;
};

# Represents cell properties.
#
# + address - Represents the range reference in A1-style. Address value will contain the Sheet reference
# (e.g. Sheet1!A1:B4)  
# + addressLocal - Represents cell reference in the language of the user  
# + columnIndex - Represents the column number of the first cell in the range. Zero-indexed
# + formulas - Represents the formula in A1-style notation
# + formulasLocal - Represents the formula in A1-style notation, in the user's language and number-formatting locale. 
# For example, the English "=SUM(A1, 1.5)" formula would become "=SUMME(A1; 1,5)" in German
# + formulasR1C1 - Represents the formula in R1C1-style notation
# + hidden - Represents if cell is hidden  
# + numberFormat - Excel's number format code for the given cell 
# + rowIndex - Returns the row number of the first cell in the range. Zero-indexed
# + text - Text values of the specified range. The Text value will not depend on the cell width. The # sign substitution
# that happens in Excel UI will not affect the text value returned by the API
# + valueTypes - Represents the type of data of each cell. The data returned could be of type string, number, or a 
# boolean. Cell that contain an error will return the error string
# + values - Raw value of the specified cell
@display {label: "Cell"}
public type Cell record {
    string address;
    string addressLocal;
    int columnIndex;
    json formulas;
    json formulasLocal;
    json formulasR1C1;
    boolean hidden;
    json numberFormat;
    int rowIndex;
    json text;
    json valueTypes;
    json[][] values;
};

# Represents the Excel application that manages the workbook.
#
# + calculationMode - Returns the calculation mode used in the workbook
@display {label: "Workbook Application"}
public type WorkbookApplication record {
    string calculationMode;
};

# Represents an Excel table.
#
# + id - Returns a value that uniquely identifies the table in a given workbook. The value of the identifier remains the
# same even when the table is renamed. This property should be interpreted as an opaque string value and 
# should not be parsed to any other type 
# + name - Name of the table
# + showHeaders - Indicates whether the header row is visible or not. This value can be set to show or remove the header
# row
# + showTotals - Indicates whether the total row is visible or not. This value can be set to show or remove the total 
# row.
# + style - Constant value that represents the Table style. The possible values are: TableStyleLight1 thru 
# TableStyleLight21, TableStyleMedium1 thru TableStyleMedium28, TableStyleStyleDark1 thru TableStyleStyleDark11. A custom user-defined style present in the workbook can also be specified.
# + highlightFirstColumn - Indicates whether the first column contains special formatting
# + highlightLastColumn - Indicates whether the last column contains special formatting
# + showBandedColumns - Indicates whether the columns show banded formatting in which odd columns are highlighted
# differently from even ones to make reading the table easier
# + showBandedRows - Indicates whether the rows show banded formatting in which odd rows are highlighted differently
# from even ones to make reading the table easier
# + showFilterButton - Indicates whether the filter buttons are visible at the top of each column header. Setting this
# is only allowed if the table contains a header row
# + legacyId - Legacy Id used in older Excle clients. The value of the identifier remains the same even when the table
# is renamed. This property should be interpreted as an opaque string value and should not be parsed to
# any other type
@display {label: "Table"}
public type Table record {
    string & readonly id?;
    @display {label: "Table Name"}
    string name?;
    @display {label: "Show Headers?"}
    boolean showHeaders?;
    @display {label: "Show Totals"}
    boolean showTotals?;
    @display {label: "Table Style"}
    string style?;
    @display {label: "Highlight First Column?"}
    boolean highlightFirstColumn?;
    @display {label: "Highlight Last Column?"}
    boolean highlightLastColumn?;
    @display {label: "Show Banded Columns?"}
    boolean showBandedColumns?;
    @display {label: "Show Banded Rows?"}    
    boolean showBandedRows?;
    @display {label: "Show Filter Button?"}
    boolean showFilterButton?;
    string & readonly legacyId?;
};

# Represents row properties.
#
# + index - Returns the index number of the row within the rows collection of the table. Zero-indexed
# + values - Represents the raw values of the specified range. The data returned could be of type string, number, or a 
# boolean. Cell that contain an error will return the error strings
@display {label: "Row"}
public type Row record {
    int index;
    json[][] values;
};

# Chart object in a workbook.
#
# + id - Chart ID
# + height - The height, in points, of the chart object
# + left - The distance, in points, from the left side of the chart to the worksheet origin
# + name - The name of a chart
# + top - The distance, in points, from the top edge of the object to the top of row 1 (on a worksheet) or the top of 
# the chart area (on a chart)
# + width - The width, in points, of the chart object
@display {label: "Chart"}
public type Chart record {
    string & readonly id?;
    @display {label: "Chart Height"}
    float height?;
    @display {label: "Distance from Left"}
    float left?;
    @display {label: "Chart Name"}
    string name?;
    @display {label: "Distance from Top"}
    float top?;
    @display {label: "Chart Width"}
    float width?;
};

# Represents column properties.
#
# + id - A unique key that identifies the column within the table. This property should be interpreted as an opaque 
# string value and should not be parsed to any other type
# + name - The name of the table column
# + index - The index number of the column within the columns collection of the table
# + values - Raw values of the specified range. The data returned could be of type string, number, or a
# boolean. Cell that contain an error will return the error string
@display {label: "Column"}
public type Column record {
    string id;
    string name?;
    int index;
    json[][] values;
};

# Specifies the calculation type to use in the workbook.
@display {label: "Calculation Type"}
public enum CalculationType {
    RECALCULATE = "Recalculate",
    FULL = "Full",
    FULL_REBUILD = "FullRebuild"
}

# Specifies the way columns or rows are used as data series on the chart.
@display {label: "Series By"} 
public enum SeriesBy {
    AUTO = "Auto",
    BY_COLUMNS = "Columns",
    BY_ROWS = "Rows"
}

# Specifies Visibility options in the worksheet.
@display {label: "Visibility"} 
public enum Visibility {
    VISIBLE = "Visible",
    HIDDEN = "Hidden",
    VERY_HIDDEN = "VeryHidden"
}

# Specifies the options used to scale the chart to the specified dimensions.
@display {label: "Chart Fitting Mode"} 
public enum FittingMode {
    FIT = "Fit",
    FIT_AND_CENTER = "FitAndCenter",
    FILL = "Fill"
}
