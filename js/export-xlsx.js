(function(H) {
    // Returns the first value that is not null or undefined.
    var pick = H.pick;

    // This extends the getDataRows function to only include rows for points visible in
    // the current chart view. Author: Torstein Honsi
    // Source: https://github.com/highcharts/highcharts/issues/7913#issuecomment-371052869
    H.wrap(H.Chart.prototype, 'getDataRows', function(proceed, multiLevelHeaders) {
        var rows = proceed.call(this, multiLevelHeaders),
            xMin = this.xAxis[0].min,
            xMax = this.xAxis[0].max;
        rows = rows.filter(function(row) {
            return typeof row.x !== 'number' || (row.x >= xMin && row.x <= xMax);
        });
        return rows;
    });

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


    // Add default XLSX exporting options to Chart init function
    H.wrap(H.Chart.prototype, 'init', function(proceed) {
        var exporting = arguments[1].exporting;
        if (exporting) {
            exporting.xlsx = exporting.xlsx || {};
            exporting.xlsx.worksheet = exporting.xlsx.worksheet || {};
            exporting.xlsx.worksheet.autoFitColumns = exporting.xlsx.worksheet.autoFitColumns || false;
            exporting.xlsx.worksheet.sheetName = exporting.xlsx.worksheet.sheetName || 'Sheet1';
            exporting.xlsx.worksheet.categoryColumn = exporting.xlsx.worksheet.categoryColumn || {};
            exporting.xlsx.worksheet.headerStyle = exporting.xlsx.worksheet.headerStyle || {};
            exporting.xlsx.workbook = exporting.xlsx.workbook || {};
            exporting.xlsx.workbook.fileProperties = exporting.xlsx.workbook.fileProperties || {};
            exporting.xlsx = deepClone(exporting.xlsx);
        }
        proceed.apply(this, Array.prototype.slice.call(arguments, 1));
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

    H.Chart.prototype.downloadXLSX = function() {
        // Need to add 0.71 to desired column width for Calibri 11pt font to get the width
        // of the column in the exported file to match the desired width.
        // https://github.com/exceljs/exceljs/issues/744
        // Excel's default column width for Calibri 11pt font is 8.43 characters
        var DEFAULT_COL_WIDTH = 8.43 + 0.71;

        // Highcharts default csv export date format as an Excel date format
        var DEFAULT_DATE_FORMAT = 'yyyy-mm-dd hh:mm:ss';

        var chart = this;

        // Array of chart data rows with header row removed
        var dataRows = chart.getDataRows().slice(1);

        // Store xlsx exporting options
        var xlsxOptions = this.options.exporting.xlsx;

        // Set export worksheet name to options.exporting.xlsx.worksheet.name or a
        // default of 'Sheet1'. Excel worksheet name length cannot exceed 31 characters
        var worksheetName = pick(xlsxOptions.worksheet.sheetName.substring(0, 31), 'Sheet1');

        // Get category column Excel number format
        var exportCategoryFormat;

        if (chart.axes[0].isDatetimeAxis) {
            exportCategoryFormat = xlsxOptions.worksheet.categoryColumn.numberFormat || DEFAULT_DATE_FORMAT;
        } else {
            exportCategoryFormat = xlsxOptions.worksheet.categoryColumn.numberFormat || null;
        }

        // Initialize an empty workbook and worksheet
        var workbook = new ExcelJS.Workbook();
        var worksheet = workbook.addWorksheet(worksheetName);

        // Array with column header titles as they will appear when exported
        var columnHeaders = [];
        // Array that will contain exceljs column objects
        var worksheetColumns = [];

        // Add chart category column to worksheet columns array
        var categoryColumnOptions = {}

        // Use user-specified column header or default category column name
        var categoryColumnHeader = xlsxOptions.worksheet.categoryColumn.title || chart.getDataRows()[0][0];
        columnHeaders.push(categoryColumnHeader);
        categoryColumnOptions.header = categoryColumnHeader;
        categoryColumnOptions.key = categoryColumnHeader;

        if (exportCategoryFormat) {
            categoryColumnOptions.style = { numFmt: exportCategoryFormat };
            // Have to set a column width if column style is applied or the column will
            // not appear in the exported file.
            // https://github.com/exceljs/exceljs/issues/458
            categoryColumnOptions.width = DEFAULT_COL_WIDTH;
            // This width is overridden if autoFitColumns is true
        }
        worksheetColumns.push(categoryColumnOptions);

        // Add each chart series to the worksheet columns array
        chart.series.forEach(function(series, index) {
            var seriesOptions = series.options;
            var seriesColumnOptions = {};

            // Use user-specified column header or default series name
            var seriesColumnHeader = seriesOptions.xlsx.name || seriesOptions.name;
            columnHeaders.push(seriesColumnHeader);

            seriesColumnOptions.header = seriesColumnHeader;
            seriesColumnOptions.key = seriesColumnHeader;
            // var columnOptions = { header: seriesColumnHeader, key: seriesColumnHeader }

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
        });

        // Add all columns to the worksheet object
        worksheet.columns = worksheetColumns;

        // Set date values in category column for correct Excel export if chart has
        // a datetime axis
        if (chart.axes[0].isDatetimeAxis) {
            dataRows.forEach(function(values, index) {
                // Need to add 'Z' to make sure values are in UTC time
                var jsDate = new Date(values[0] + 'Z');
                dataRows[index][0] = jsDate;
            });
        }

        // Add the data to the worksheet
        worksheet.addRows(dataRows);

        // If enabled, autofit columns by setting column widths to the width of the
        // cell with the most characters. This requires the SSF module to be loaded.
        // https://github.com/SheetJS/ssf/blob/master/ssf.js
        // https://cdn.statically.io/gh/SheetJS/ssf/e267d1d6/ssf.js
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
