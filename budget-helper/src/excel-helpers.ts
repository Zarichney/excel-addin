
import { Observable } from 'rxjs';

export function NamedRangeValues$(rangeName: string): Observable<string> {
    return new Observable(observer => {
        Excel.run(async (context) => {
            const namedItem = context.workbook.names.getItem(rangeName);
            const range = namedItem.getRange().getUsedRange();
            range.load('values');

            try {
                await context.sync();
            } catch (error) {
                console.error("Error getting range named '" + rangeName + "':", error);
                throw error; // Rethrow to ensure calling code is aware of the failure
            }

            if (range.values && range.values.length > 0 && range.values[0].length > 0) {
                // Flatten the 2D array to 1D and filter out empty values
                const values = range.values.flat().filter((value, index, self) => value !== "" && self.indexOf(value) === index);

                values.forEach(value => observer.next(value));
            } else {
                observer.next('Named range is empty or does not exist');
            }

            observer.complete();
        }).catch(error => observer.error(error));
    });
}

export function TableRows$(tableName: string): Observable<any> {
    return new Observable(observer => {
        Excel.run(async (context) => {
            const table = context.workbook.tables.getItem(tableName);
            const headerRow = table.getHeaderRowRange().load('values');
            const dataRows = table.getDataBodyRange().load('values');

            try {
                await context.sync();
            } catch (error) {
                console.error("Error getting rows for table '" + tableName + "':", error);
                throw error; // Rethrow to ensure calling code is aware of the failure
            }

            const headers = headerRow.values[0];
            const rows = dataRows.values;

            rows.forEach(row => {
                const rowObject = headers.reduce((obj, header, index) => {
                    obj[header] = row[index];
                    return obj;
                }, {});

                observer.next(rowObject);
            });

            observer.complete();
        }).catch(error => observer.error(error));
    });
}

export async function AddToTable(tableName: string, data: any) {
    // Turn any object into an array of values
    var row = Object.values(data);
    WriteToTable(tableName, [row]);
}

export async function WriteToTable(tableName: string, data: any[]) {
    try {

        if (data.length === 0) {
            console.warn("No data to write to table '" + tableName + "'.");
            return;
        }

        await Excel.run(async (context) => {

            const table = context.workbook.tables.getItemOrNullObject(tableName);
            await context.sync(); // Ensure table is loaded or null if not found

            if (table.isNullObject) {
                throw new Error(`Table "${tableName}" not found.`);
            }

            const addedRows = table.rows.add(null, data);

            // Load the address of the added rows to access them later for formatting
            // addedRows.load("address");

            await context.sync();

            // Example: Set number format for the first column of the added rows
            // const range = context.workbook.worksheets.getActiveWorksheet().getRange(addedRows.address);
            // range.numberFormat = [[null, "mm-dd-yyyy", null]]; // Assume the second column needs date formatting
            // Adjust the numberFormat array to match the formatting requirements of your table columns
            // await context.sync();
        });
    } catch (error) {
        console.error("Error writing to table:", error);
        console.error("Could not add data to table '" + tableName + "':", data);
        throw error; // Rethrow to ensure calling code is aware of the failure
    }
}

export async function UpdateTableRow(tableName: string, rowIndex: number, data: any) {
    await Excel.run(async (context) => {
        const table = context.workbook.tables.getItem(tableName);
        // Assuming row index is based on the data body range (not including the header)
        const rowRange = table.getDataBodyRange().getRow(rowIndex);
        // Convert the entry object to an array of values based on table headers
        const headers = table.getHeaderRowRange().load('values');
        await context.sync(); // Load headers

        const updatedValues = headers.values[0].map(header => data[header] ?? null);
        rowRange.values = [updatedValues];
        await context.sync();
    }).catch(error => {
        console.error("Error updating row in table:", error);
        throw error; // Rethrow to ensure calling code is aware of the failure
    });
}

