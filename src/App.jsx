import { useState } from 'react'
import XLSX from 'xlsx';
import { HotTable } from '@handsontable/react';
import 'handsontable/dist/handsontable.full.css';

import './App.css'

// 解析 Excel 文件中的所有合并单元格
function getAllMergedCells(sheet) {
  const mergeCells = [];
  const merges = sheet['!merges'];
  
  if (merges) {
    merges.forEach(({ s, e }) => {
      mergeCells.push({
        row: s.r,
        col: s.c,
        rowspan: e.r - s.r + 1,
        colspan: e.c - s.c + 1,
      });
    });
  }
  console.log(mergeCells);
  return mergeCells;
}

function App() {
  const [data, setData] = useState([]);
  const [mergedCells, setMergedCells] = useState([]);
  const handleFileUpload = (e) => {
    console.log();
    const { files } = e.target
    if(files.length == 0) return;
    const reader = new FileReader();

    reader.onload = (e) => {
      const workbook = XLSX.read(e.target.result, { type: 'binary' });

      const sheetName = workbook.SheetNames[0];
      const worksheet = workbook.Sheets[sheetName];
      const sheetData = XLSX.utils.sheet_to_json(worksheet, {header: 1, range: 0});
      console.log(sheetData);
      // 解析合并单元格信息
      const mergedCells = getAllMergedCells(worksheet);

      setMergedCells(mergedCells);
      setData(sheetData);
    };

    reader.readAsBinaryString(files[0]);
  }

  const handleFileDownload = () => {
    const worksheet = XLSX.utils.aoa_to_sheet(data);
    const newWorkbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(newWorkbook, worksheet, 'Sheet1');
    XLSX.writeFile(newWorkbook, 'output.xlsx');
  };

  return (
    <>
      <input type="file" accept=".xlsx,.xls" onChange={handleFileUpload} />
      <button onClick={handleFileDownload}>Download</button>
      <div style={{height: 500, width: 400}}>
        <HotTable
          data={data} 
          colHeaders={true} 
          rowHeaders={true} 
          mergeCells={mergedCells}
          licenseKey="non-commercial-and-evaluation"  
        />
      </div>


    </>
  )
}

export default App
