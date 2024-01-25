import { Component } from '@angular/core';
import * as XLSX from 'xlsx';
import { ExcelInterface } from './interfaces/excelInterface';

@Component({
  selector: 'app-root',
  templateUrl: './app.component.html',
  styleUrls: ['./app.component.css'],
})
export class AppComponent {
  title = 'excel-exercise';
  ExcelData: any;
  valor = '';
  constructor() {}

  ReadExcel(event: any) {
    let file = event.target.files[0];
    let fileName = file.name;

    let extension = fileName.slice(0, fileName.indexOf('.'));

    console.log(extension);

    if (file.name != 'comisiones_masivas.xlsx') {
      alert('N de arc');
      this.ExcelData = [];
      this.valor = '';
      return;
    } else {
      let fileReader = new FileReader();
      fileReader.readAsBinaryString(file);

      fileReader.onload = (e) => {
        var workBook = XLSX.read(fileReader.result, { type: 'binary' });
        var sheetNames = workBook.SheetNames;
        this.ExcelData = XLSX.utils.sheet_to_json(
          workBook.Sheets[sheetNames[0]]
        );

        console.log(this.ExcelData);
      };
    }
  }

  multiplicar(i: number) {
    let resultado = i + 1;
  }

  pagar() {
    this.ExcelData.forEach((i: any) => {
      if (i.Validacion < 1) {
        i.estado = 'Procesado';
        console.log(this.ExcelData);
      } else {
        i.estado = 'Procesado con error';
        console.log(this.ExcelData);
      }
    });
  }
}