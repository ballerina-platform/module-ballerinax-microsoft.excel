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

# Reprents worksheet properties
#
# + id - Worksheet id  
# + position - The zero-based position of the worksheet within the workbook
# + name - Worksheet name
# + visibility - Worksheet visibilty
public type Worksheet record {
    string id?;
    int position?;
    string name?;
    string visibility?;
};

# Represents cell properties.
#
# + address - Reference in A1-style  
# + addressLocal - Reference in the language of the user  
# + columnIndex - Column number of the cell
# + formulas - Represents the formula in A1-style notation
# + formulasLocal - Represents the formula in A1-style notation and in the user's language
# + formulasR1C1 - Represents the formula in R1C1-style notation
# + hidden - Represents if cell is hidden  
# + numberFormat - Excel's number format code for the given cell 
# + rowIndex - Row number of the cell
# + text - Text values of the specified cell 
# + valueTypes - The type of data of each cell
# + values - Raw value of the specified cell
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
    json values;
};

# Represents the Excel application that manages the workbook.
#
# + calculationMode - The calculation type to use  
public type Application record {
    string calculationMode;
};

# Table configuration
#
# + address - Address or name of the range object representing the data source
# + hasHeaders - Boolean value that indicates whether the data being imported has column labels
public type TableConfiguration record {
    string address;
    string hasHeaders?;
};

# Represents an Excel table.
#
# + id - Table id  
# + name - Table name
# + showHeaders - Indicates whether the header row is visible or not
# + showTotals - Indicates whether the total row is visible or not
# + style - Constant value that represents the Table style
# + highlightFirstColumn - Indicates whether the first column contains special formatting
# + highlightLastColumn - Indicates whether the last column contains special formatting
# + showBandedColumns - Indicates whether the columns show banded formatting
# + showBandedRows - Indicates whether the rows show banded formatting
# + showFilterButton - Indicates whether the filter buttons are visible at the top of each column header
# + legacyId - Used in older Excle clients. The value of the identifier remains the same even when the table is renamed
public type Table record {
    string id?;
    string name?;
    boolean showHeaders?;
    boolean showTotals?;
    string style?;
    boolean highlightFirstColumn?;
    boolean highlightLastColumn?;
    boolean showBandedColumns?;
    boolean showBandedRows?;
    boolean showFilterButton?;
    string legacyId?;
};

# Row properties.
#
# + index - Specifies the relative position of the new row
# + values - A 2-dimensional array of unformatted values of the table rows
public type Row record {
    int index?;
    json values;
};

# Chart object in a workbook.
#
# + id - Chart id
# + height - The height, in points, of the chart object
# + left - The distance, in points, from the left side of the chart to the worksheet origin
# + name - Chart name
# + top - The distance, in points, from the top edge of the object to the top of row 1 (on a worksheet) or the top of 
# the chart area (on a chart)
# + width - The width, in points, of the chart object
public type Chart record {
    string id?;
    float height?;
    float left?;
    string name?;
    float top?;
    float width?;
};

# Configuration for a new chart.
#
# + 'type - the type of a chart  
public type ChartConfiguration record {
    string 'type?;
    *Data;
};

# Chart Data.
#
# + sourceData - The Range object corresponding to the source data
# + seriesBy - Specifies the way columns or rows are used as data series on the chart
public type Data record {
    json sourceData;
    SeriesBy seriesBy?;
};

# Column properties.
#
# + id - Column id
# + name - Column name
# + index - The index number of the column within the columns collection of the table
# + values - Column values
public type Column record {
    string id?;
    string name?;
    int index?;
    json values;
};

# Chart image.
#
# + value - Image in base-64 string
public type ChartImage record {
    string value;
};

# Chart position.
#
# + startCell - The start cell. This is where the chart will be moved to
# + endCell - The end cell. If specified, the chart's width and height will be set to fully cover up this cell/range
public type ChartPosition record {
    string startCell;
    string endCell?;
};

# Query options.
#
# + count - Retrieve the total count of matching resources
# + expand - Retrieve related resources
# + filter - Filter results (rows)
# + format - Return the results in the specified media format
# + orderBy - Order results
# + search - Return results based on search criteria
# + 'select - Filter properties (columns)
# + skip - Set the number of items to skip at the start of a collection
# + top - Set the page size of results
public type Query record {
    boolean count?;
    string expand?;
    string filter?;
    string format?;
    string orderBy?;
    string search?;
    string 'select?;
    int skip?;
    int top?;
};

public enum CalculationType {
    RECALCULATE = "Recalculate",
    FULL = "Full",
    FULL_REBUILD = "FullRebuild"
}

public enum SeriesBy {
    AUTO = "Auto",
    BY_COLUMNS = "Columns",
    BY_ROWS = "Rows"
}

public enum Shift {
    DOWN = "down",
    RIGHT = "right"
}

public enum Visibility {
    VISIBLE = "Visible",
    HIDDEN = "Hidden",
    VERY_HIDDEN = "VeryHidden"
}
