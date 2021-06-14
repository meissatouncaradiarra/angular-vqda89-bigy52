import { Component, OnInit, ViewChild } from '@angular/core';
import { data, employeeData } from './data';
import { GridComponent, ToolbarItems, ExcelExportProperties } from '@syncfusion/ej2-angular-grids';
import { ClickEventArgs } from '@syncfusion/ej2-angular-navigations';
import { Workbook } from '@syncfusion/ej2-excel-export';

@Component({
  selector: 'app-root',
  templateUrl: 'app.component.html'
})
export class AppComponent {
  public fData: object[];
  public objGrid: any = [];
  public sData: object[];
  public toolbarOptions: ToolbarItems[];
  @ViewChild('grid1') public fGrid: GridComponent;
  @ViewChild('grid2') public sGrid: GridComponent;

  ngOnInit(): void {
    this.fData = data.slice(0, 5);
    this.sData = employeeData.slice(0, 5);
    this.toolbarOptions = ['ExcelExport'];
  }

  toolbarClick = (args: ClickEventArgs) => {
    var names = ["OrderDetail", "EmployeeDetail"];
    for (var i = 0; i < document.querySelectorAll(".e-grid").length; i++) {  // you can find all grid controls using this. 
      var grid = (document.getElementById(document.querySelectorAll(".e-grid")[i].id) as any).ej2_instances[0];
      this.objGrid.push(grid);
    }
    if (args.item.id === 'FirstGrid_excelexport') { // 'Grid_excelexport' -> Grid component id + _ + toolbar item name
      var exportData;
      const appendExcelExportProperties: ExcelExportProperties = { multipleExport: { type: 'NewSheet' } };
      if (this.objGrid.length > 1) {  
        var firstGridExport = this.objGrid[0].excelExport(appendExcelExportProperties, true).then(function (fData) {
          fData.worksheets[0].name = names[0];
          exportData = fData;
          for (var j = 1; j < this.objGrid.length - 1; j++) {    // iterate grids here 
            this.objGrid[j].excelExport(appendExcelExportProperties, true, exportData).then(function (wb) {
              exportData = wb;
              if (exportData.worksheets.length === (this.objGrid.length - 1)) {
                for (var k = 0; k < exportData.worksheets.length; k++) {
                  if (!exportData.worksheets[k].name) {
                    exportData.worksheets[k].name = names[k];
                  }
                }
              }
            })
          }
          var lastGridExport = this.objGrid[this.objGrid.length - 1].excelExport(appendExcelExportProperties, true, exportData).then(function (wb) {
            wb.worksheets[wb.worksheets.length - 1].name = names[this.objGrid.length - 1];
            const book = new Workbook(wb, 'xlsx');
            book.save('Export.xlsx');
          }.bind(this));
        }.bind(this));
      }
    }
  }
}