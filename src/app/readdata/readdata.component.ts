import { NON_BINDABLE_ATTR } from '@angular/compiler/src/render3/view/util';
import { Component, OnInit } from '@angular/core';
import * as Excel from 'exceljs/dist/exceljs.min.js';
import * as XLSX from 'xlsx';
import * as fs from 'file-saver'

@Component({
  selector: 'app-readdata',
  templateUrl: './readdata.component.html',
  styleUrls: ['./readdata.component.css']
})
export class ReaddataComponent implements OnInit {
  constructor() { }
  data:any[]=[
    ['name','age','address'],
    ['Nhan','18','asdvcxbrf86qwe4r1'],
  ];
  ngOnInit(): void {
  }
  readExcel($e){
      const target:DataTransfer=<DataTransfer>($e.target);
      const workbook=new Excel.Workbook();
      const reader=new FileReader();
      reader.readAsArrayBuffer(target.files[0]);
      reader.onload=()=>{
        this.data=[];
          const buffer=reader.result;
          workbook.xlsx.load(buffer).then(workbook=>{
              //console.log(workbook,'workbook instance');
              workbook.eachSheet((sheet,id)=>{
                  const datasm=[];
                  sheet.eachRow(row=>{
                    this.data.push(row.values);
                  })
              })
          })
          console.log(this.data);
      }
      
  }
  writingtoExcel():void{
      const workbook=new Excel.Workbook();
      const worksheet=workbook.addWorksheet('My Sheet');
      for(let x1 of this.data)
      {
        let x2=Object.keys(x1);
        let temp=[];
        for(let y of x2)
        {  
          temp.push(x1[y])
        }
        worksheet.addRow(temp);
      }
      let fname="Data-Excel";
      workbook.xlsx.writeBuffer().then((data) => {
        let blob = new Blob([data], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
        fs.saveAs(blob, fname+'-'+new Date().valueOf()+'.xlsx');
      });
  }
}
