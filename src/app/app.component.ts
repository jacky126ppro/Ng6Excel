import { Component } from '@angular/core';
import * as XLSX from 'xlsx';


@Component({
  selector: 'app-root',
  templateUrl: './app.component.html',
  styleUrls: ['./app.component.css']
})
export class AppComponent {
  dropzoneStyle: any = null;
  reader: FileReader = new FileReader();
  files: FileList = null;
  workBook: XLSX.WorkBook = null;
  workSheetNames: string[] = [];
  workSheets: { [sheet: string]: XLSX.WorkSheet } = null;
  currentWorkSheet: XLSX.WorkSheet = null;
  currentData: any[] = [];
  constructor() {
    this.reader.onload = this.setReaderOnloadHandler();
  }

  readFile($event) {
    let files = $event.target.files;
    this.files = files;
    console.log(files);
    this.dropzoneStyle = { borderColor: 'green' };

    // 取第一個檔
    let file = files[0];

    this.reader.readAsBinaryString(file);
  }
  private setReaderOnloadHandler() {
    return () => {
      let data = this.reader.result;
      this.workBook = XLSX.read(data, { type: 'binary' });
      this.workSheetNames = this.workBook.SheetNames;
      this.workSheets = this.workBook.Sheets;

      for (let i = 0; i < this.workSheetNames.length; i++) {
        this.currentWorkSheet = this.workSheets[this.workSheetNames[i]];
        this.currentData = this.currentData.concat(XLSX.utils.sheet_to_json(this.workSheets[this.workSheetNames[i]]));
      }
    };
  }
}
