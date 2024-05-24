import { Component, OnInit, Input, Output, EventEmitter } from '@angular/core';



@Component({
    selector: 'app-root',
    templateUrl: './app.component.html',
    styleUrls: ['./app.component.css']
})
export class AppComponent implements OnInit {
    @Input() sheets: Array<string> = [];
    sheetSelected: string = '';

    selectSheet(sheetName: any): void {
        this.sheetSelected = sheetName.target.value as string;

        Excel.run(async (context) => {
            let sheet = context.workbook.worksheets.getItem(this.sheetSelected);
            sheet.activate();
            sheet.load("name");

            await context.sync();
            console.log(`The active worksheet is "${sheet.name}"`);
        });
    }

    ngOnInit() {
        Office.onReady((info: any) => {
            if (info.host === Office.HostType.Excel) {
                this.readExcelData();
            }
        });
    }

    async readExcelData() {
        let sheets: Array<string> = [];
        Excel.run(async (context) => {
            const sheetss = context.workbook.worksheets;
            sheetss.load('items/name');

            await context.sync();

            if (sheetss.items.length > 1) {
                console.log(`There are ${sheetss.items.length} worksheets in the workbook:`);
            } else {
                console.log('There is one worksheet in the workbook:');
            }

            sheets = sheetss.items.map(sheet => sheet.name);
            await context.sync();
            this.sheets = sheets;
        });
    };

    async getColumnNames() {
        return Excel.run(async (context) => {
            let sheet = context.workbook.worksheets.getItem(this.sheetSelected);
            let range = sheet.getUsedRange();
            range.load("values");
            await context.sync();
            const columnNames: string[] = range.values[0];
            console.log("Column names:", columnNames);
            return columnNames;
        });
    }

    async getDataByColumn(columnName: string) {
        return Excel.run(async (context) => {
            const sheet = context.workbook.worksheets.getActiveWorksheet();
            const usedRange = sheet.getUsedRange();
            usedRange.load("values");

            await context.sync();

            const values: any[][] = usedRange.values;

            const columnIndex = values[0].indexOf(columnName);
            if (columnIndex === -1) {
                throw new Error(`Column '${columnName}' not found.`);
            }

            const data: any[] = [];
            for (let i = 1; i < values.length; i++) {
                const cellValue = values[i][columnIndex];
                data.push(cellValue);
            }

            return data;
        });
    }

    async getDataByExistingColumns() {
        try {
            const columnNames = await this.getColumnNames();
            console.log("Column names:", columnNames);

            for (const columnName of columnNames) {
                const data = await this.getDataByColumn(columnName);
                console.log(`Data for column '${columnName}':`, data);
            }
        } catch (error) {
            console.error("Error:", error);
        }
    }
}
