(function(H) {
    // Returns the first value that is not null or undefined.
    var pick = H.pick;

    // Add "Download XLSX" to the exporting menu in place of "Download XLS". Source:
    // https://jsfiddle.net/gh/get/library/pure/highcharts/highcharts/tree/master/samples/highcharts/export-data/xlsx/
    H.getOptions().lang.downloadXLSX = 'Download XLSX';

    // Add the menu item handler
    H.getOptions().exporting.menuItemDefinitions.downloadXLSX = {
        textKey: 'downloadXLSX',
        onclick: function () {
            this.downloadXLSX();
        }
    };

    // Replace the menu item
    var menuItems = H.getOptions().exporting.buttons.contextButton.menuItems;
    menuItems[menuItems.indexOf('downloadXLS')] = 'downloadXLSX';

    // Moved initalization of xlsx export options to addEvent so they are defined on
    // chart load and hence can be referenced in the chart load function
    H.addEvent(H.Chart, 'load', function(e) {
        var chart = e.target;
        // Check if all series are pie series. Need to disable filtered getDataRows function
        // for pie series or only the first row is exported.
        // See https://github.com/chris-vecchio/export-xlsx-exceljs/issues/2
        var allPieSeries = false;

        var seriesTypes = [];
        chart.series.forEach(function(series, index) {
            seriesTypes.push(series.type);
        });

        // Remove duplicates and convert back to array
        seriesTypes = Array.from(new Set(seriesTypes));

        // If there is only one series type and it is 'pie' series type
        if (seriesTypes.length == 1 && seriesTypes[0] == 'pie') allPieSeries = true;

        var exporting = chart.options.exporting;
        if (exporting) {
            exporting.xlsx = exporting.xlsx || {};
            // worksheet options
            exporting.xlsx.worksheet = exporting.xlsx.worksheet || {};
            exporting.xlsx.worksheet.autoFitColumns = exporting.xlsx.worksheet.autoFitColumns || false;
            exporting.xlsx.worksheet.sheetName = exporting.xlsx.worksheet.sheetName || 'Sheet1';
            exporting.xlsx.worksheet.categoryColumn = exporting.xlsx.worksheet.categoryColumn || {};
            exporting.xlsx.worksheet.headerStyle = exporting.xlsx.worksheet.headerStyle || {};
            // workbook options
            exporting.xlsx.workbook = exporting.xlsx.workbook || {};
            exporting.xlsx.workbook.fileProperties = exporting.xlsx.workbook.fileProperties || {};
            exporting.xlsx = deepClone(exporting.xlsx);
            // check for all pie series
            chart.options.exporting.xlsx.allPieSeries = allPieSeries;
        }
    });

    H.wrap(H.Chart.prototype, 'getDataRows', function(proceed, multiLevelHeaders) {
        var rows = proceed.call(this, multiLevelHeaders),
            xMin = this.xAxis[0].min,
            xMax = this.xAxis[0].max;
        // Only filter rows if all series are not 'pie' series
        // See https://github.com/chris-vecchio/export-xlsx-exceljs/issues/2
        if (this.options.exporting.xlsx.allPieSeries == false) {
            rows = rows.filter(function(row) {
                return typeof row.x !== 'number' || (row.x >= xMin && row.x <= xMax);
            });
        }
        return rows;
    });

    // Add default XLSX exporting options to Series init function
    H.wrap(H.Series.prototype, 'init', function(proceed, chart, options) {
        options.xlsx = options.xlsx || {};
        options.xlsx.numberFormat = options.xlsx.numberFormat || undefined;
        options.xlsx.name = options.xlsx.name || undefined;
        options.xlsx = deepClone(options.xlsx);
        proceed.apply(this, Array.prototype.slice.call(arguments, 1));
    });

    // Used to remove blank attributes from XLSX exporting options objects
    // https://github.com/blacklabel/grouped_categories/blob/master/grouped-categories.js#L42
    function deepClone(thing) {
        return JSON.parse(JSON.stringify(thing));
    }

    // exceljs requires ARGB strings for colors
    function hexToARGB(hex) {
        // Remove leading '#' if it exists and convert to hex integer
        var hex = hex.replace(/^#/, '');
        var hexString = 'FF' + hex;
        return hexString.toUpperCase();
    }

    // Convert string character width to Excel column width
    // Source: https://github.com/SheetJS/sheetjs/blob/master/bits/46_stycommon.js
    function char2width(chr) {
        // Excel Max Digit Width
        // Info: https://github.com/SheetJS/sheetjs/blob/master/docbits/62_colrow.md
        var MDW = 6;
        return Math.ceil((Math.round((chr * MDW + 5) / MDW * 256)) / 256);
    }

    // Check if object is empty
    // https://stackoverflow.com/a/50210676/6579114
    function isEmpty(obj) {
      return !obj || Object.keys(obj).length === 0;
    }

    // Parse a string date in format yyyy-mm-dd hh:mm:ss to a date object that exports
    // correctly to Excel
    var parseDateStrToExcel = function(d, fixmonth) {
        if (fixmonth === undefined) fixmonth = true;
        if (d == '') {return null;}

        var datearr = d.split(' ');
        var thisDate = datearr[0].split('-');
        var thisTime = datearr[1].split(':');
        var year = thisDate[0];
        var month = thisDate[1];
        var day = thisDate[2];
        var hours = thisTime[0];
        var minutes = thisTime[1];
        var seconds = thisTime[2];

        // Subtract 1 from month if month from data is not zero indexed.
        if (fixmonth == true) {
            month = month - 1;
        }

        var jsDate = new Date(year, month, day);

        // Need to set the time parts after creating the date to get correct format
        // in Excel
        jsDate.setUTCHours(hours);
        jsDate.setUTCMinutes(minutes);
        jsDate.setUTCSeconds(seconds);

        return jsDate;
    };

    // Need to add 0.71 to desired column width for Calibri 11pt font to get the width
    // of the column in the exported file to match the desired width.
    // https://github.com/exceljs/exceljs/issues/744
    // Excel's default column width for Calibri 11pt font is 8.43 characters
    var DEFAULT_COL_WIDTH = 8.43 + 0.71;

    // Highcharts default csv export date format as an Excel date format
    var DEFAULT_DATE_FORMAT = 'yyyy-mm-dd hh:mm:ss';


    H.Chart.prototype.downloadXLSX = function() {
        var chart = this;

        var columnHeaderFormatter = chart.options.exporting.csv.columnHeaderFormatter;

        // Array of columnHeaders from getDataRows function
        var columnHeaders = chart.getDataRows()[0];

        // Boolean for datetime/non-datetime x-axis
        var xAxisIsDatetime = chart.xAxis[0].options.type == "datetime" ? true : false;

        // Array of chart data rows with header row removed
        var dataRows = chart.getDataRows().slice(1);

        // Store xlsx exporting options
        var xlsxOptions = this.options.exporting.xlsx;

        // Set export worksheet name to options.exporting.xlsx.worksheet.name or a
        // default of 'Sheet1'. Excel worksheet name length cannot exceed 31 characters
        var worksheetName = xlsxOptions.worksheet.sheetName ? xlsxOptions.worksheet.sheetName.substring(0, 31) : 'Sheet1';

        // Initialize an empty workbook and worksheet
        var workbook = new ExcelJS.Workbook();
        var worksheet = workbook.addWorksheet(worksheetName);


        /**
         * Get and set series column options
         */

        // Array that will contain exceljs column objects
        var worksheetColumns = [];

        chart.series.forEach(function(series, index) {
            var seriesOptions = series.options;
            var seriesColumnOptions = {};

            if (series.options.includeInDataExport !== false &&
                    !series.options.isInternal &&
                    series.visible !== false // #55
            ) {

                // Use user-specified xlsx column header if specified, if not defined,
                // use the result of seriesColumnHeaderFormatter if defined, and finally
                // use default series name
                var seriesColumnHeader = seriesOptions.xlsx.name ? seriesOptions.xlsx.name : columnHeaderFormatter ? columnHeaderFormatter(series) : seriesOptions.name;

                seriesColumnOptions.header = seriesColumnHeader;
                seriesColumnOptions.key = seriesColumnHeader;

                // Add number format if specified in series options
                // var columnNumFmt = seriesOptions.xlsx.numberFormat || null;
                if (seriesOptions.xlsx.numberFormat) {
                    seriesColumnOptions.style = { numFmt: seriesOptions.xlsx.numberFormat };
                    // Have to set a column width if column style is applied or the column will
                    // not appear in the exported file.
                    // https://github.com/exceljs/exceljs/issues/458
                    seriesColumnOptions.width = DEFAULT_COL_WIDTH;
                    // This width is overridden if autoFitColumns is true
                }
                worksheetColumns.push(seriesColumnOptions);
            }
        });


        /**
         * Get and set category column options
         */

        // Get category column header
        var categoryColumnHeader = columnHeaders[0];

        // Store index values for category columns. This is normally just the zeroth
        // column but can be more in combination charts. We use these indices when inserting
        // category columns into the ExcelJS worksheet columns object
        var categoryColumnHeaderIndexes = [];
        for(var i = 0; i < columnHeaders.length; i++) {
            if (columnHeaders[i] === categoryColumnHeader) {
                categoryColumnHeaderIndexes.push(i);
            }
        }

        // Category column header definition logic is as follows:
        // if defined: chart.options.exporting.xlsx.worksheet.categoryColumn.title
        // else if columnHeaderFormatter defined: the result of columnHeaderFormatter for the xAxis
        // finally: default Highcharts categoryColumnHeader
        if (xlsxOptions.worksheet.categoryColumn.title) {
            categoryColumnHeader = xlsxOptions.worksheet.categoryColumn.title;
        } else if (columnHeaderFormatter) {
            categoryColumnHeader = columnHeaderFormatter(chart.xAxis[0]);
        }

        // Category column Excel number format
        var categoryColumnNumberFormat;

        if (xAxisIsDatetime) {
            categoryColumnNumberFormat = xlsxOptions.worksheet.categoryColumn.numberFormat || DEFAULT_DATE_FORMAT;
        } else {
            categoryColumnNumberFormat = xlsxOptions.worksheet.categoryColumn.numberFormat || null;
        }

        // Define category column options object and add category columns to worksheet
        // columns array
        var categoryColumnOptions = {};

        // Set category column options that will apply to all category columns in the
        // exported workbook.
        categoryColumnOptions.header = categoryColumnHeader;
        categoryColumnOptions.key = categoryColumnHeader;

        if (categoryColumnNumberFormat) {
            categoryColumnOptions.style = { numFmt: categoryColumnNumberFormat };
            // Have to set a column width if column style is applied or the column will
            // not appear in the exported file.
            // https://github.com/exceljs/exceljs/issues/458
            categoryColumnOptions.width = DEFAULT_COL_WIDTH;
            // This width is overridden if autoFitColumns is true
        }

        // Insert the category column definitions into the appropriate indexes in the
        // worksheetColumns array
        for(i = 0; i < categoryColumnHeaderIndexes.length; i++) {
            worksheetColumns.splice(categoryColumnHeaderIndexes[i], 0, categoryColumnOptions);
        }

        // Add all columns to the worksheet object
        worksheet.columns = worksheetColumns;

        // Set date values in category column for correct Excel export if chart has
        // a datetime axis
        if (xAxisIsDatetime) {
            dataRows.forEach(function(values, index) {
                dataRows[index][0] = parseDateStrToExcel(values[0]);
            });
        }

        // For non category/date column columns, write #N/A formula for NaN number values
        dataRows.forEach(function (values, index) {
            for (var k = 1; k < values.length-1; k++) {
                var val = dataRows[index][k];
                if (typeof val === 'number') {
                    if (isNaN(val)) {
                        val = { error: '#N/A' };
                        dataRows[index][k] = val;
                    }
                }
            }
        });
        // Add the data to the worksheet
        worksheet.addRows(dataRows);

        // If enabled, autofit columns by setting column widths to the width of the
        // cell with the most characters. This requires the SSF module to be loaded.
        // https://github.com/SheetJS/ssf/blob/master/ssf.js
        // https://cdn.statically.io/gh/SheetJS/ssf/e267d1d6/ssf.js
        var wrapColumnWidth = xlsxOptions.worksheet.wrapColumnWidth || 20;

        if (xlsxOptions.worksheet.autoFitColumns === true) {
            // Get the length of the longest value in each column in Excel column width
            // units and set the column width.
            for (var col = 0; col < worksheet.columns.length; col++) {
                var column = worksheet.columns[col];
                var columnFormattedValues = [];
                column.values.forEach(function(value, rowIndex) {
                    var formattedValue;
                    // Don't attempt to format column header string. rowIndex = 1 because
                    // exceljs adds an undefined element at the start of the column values
                    // array so that the first value's row index matches excel's first row
                    // index of 1.
                    if (rowIndex == 1) {
                        formattedValue = value;
                    } else {
                        if ('numFmt' in column.style) {
                            formattedValue = SSF.format(column.style.numFmt, value);
                        } else {
                            formattedValue = value.toString();
                        }
                    }
                    columnFormattedValues.push(formattedValue);
                });

                // Determine the width of the longest cell in the column
                var longest = columnFormattedValues.reduce(function(a, b) { return a.length > b.length ? a : b; }, '');

                // Ensure columns aren't thinner than Excel's default column width
                // This is purely a style decision. Remove this check if you don't care
                // about having potentially very thin columns.
                var columnWidth = char2width(longest.length) > DEFAULT_COL_WIDTH ? char2width(longest.length) : DEFAULT_COL_WIDTH;
                
                if (col == 0 && columnWidth > wrapColumnWidth) {
                    worksheet.getCell('A1').alignment = { wrapText: true };
                    columnWidth = wrapColumnWidth;
                }

                column.width = columnWidth;
            }
        }

        // TESTING ALLOWING SPECIFICATION OF COLUMN HEADER FONT AND FILL OPTIONS
        if (!isEmpty(xlsxOptions.worksheet.headerStyle)) {
            var headerRow = worksheet.getRow(1);
            var headerFont = {};
            var headerFill = {};

            if (xlsxOptions.worksheet.headerStyle.font) {
                if (xlsxOptions.worksheet.headerStyle.font.bold) {
                    headerFont.bold = xlsxOptions.worksheet.headerStyle.font.bold;
                }
                if (xlsxOptions.worksheet.headerStyle.font.color) {
                    headerFont.color = { argb: hexToARGB(xlsxOptions.worksheet.headerStyle.font.color) };
                }
            }

            // Maybe add more options for fill like type and pattern see exceljs options
            if (xlsxOptions.worksheet.headerStyle.fill) {
                if (xlsxOptions.worksheet.headerStyle.fill.color) {
                    headerFill.fgColor = { argb: hexToARGB(xlsxOptions.worksheet.headerStyle.fill.color) };
                    headerFill.type = 'pattern';
                    headerFill.pattern = 'solid';
                }
            }

            // Iterate over all non-null cells in header row and apply formatting
            headerRow.eachCell(function(cell, colNumber) {
                if (headerFont) {
                    cell.font = headerFont;
                }
                if (headerFill) {
                    cell.fill = headerFill;
                }
            });
        }

        // Set any user specified workbook file properties. You can see the full list of
        // available properties at:
        // https://github.com/exceljs/exceljs/blob/master/lib/doc/workbook.js#L153
        if (xlsxOptions.workbook.fileProperties) {
            Object.keys(xlsxOptions.workbook.fileProperties).forEach(function(key) {
                var value = xlsxOptions.workbook.fileProperties[key];
                if (['lastPrinted','created','modified'].includes(key)) {
                    value = new Date(value);
                }
                workbook[key] = value;
            });
        }

        // Write the .xlsx file using FileSaver.js
        var filename = pick(this.options.exporting.filename, this.getFilename()) + '.xlsx';

        workbook.xlsx.writeBuffer().then(function(data) {
            var blob = new Blob([data], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
            saveAs(blob, filename);
        });
    };
}(Highcharts));