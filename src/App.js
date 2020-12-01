import './App.css';
import ExcelJS from "exceljs/dist/es5/exceljs.browser";
import saveAs from "file-saver";

function App() {
  const exportExcel = event => {
    var ExcelJSWorkbook = new ExcelJS.Workbook();
    var worksheet = ExcelJSWorkbook.addWorksheet("ExcelJS sheet");
    worksheet.columns = [
      { header: 'Id', key: 'id', width: 10 },
      { header: 'Name', key: 'name', width: 32 },
      { header: 'D.O.B.', key: 'dob', width: 10, outlineLevel: 1 }
    ];
    worksheet.addRow({id: 1, name: 'John Doe', dob: new Date(1970,1,1)});
    worksheet.addRow({id: 2, name: 'Jane Doe', dob: new Date(1965,1,7)});
    ExcelJSWorkbook.xlsx.writeBuffer().then(function(buffer) {
      saveAs(
        new Blob([buffer], { type: "application/octet-stream" }),
        `Sample.xlsx`
      );
    });
    console.log('export excel', worksheet);
  }
  return (
    <div className="App">
      <button onClick={exportExcel}>Export</button>
    </div>
  );
}

export default App;
