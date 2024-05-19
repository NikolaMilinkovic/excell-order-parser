import React, { useState } from 'react';
import ExcelJS from 'exceljs';
import './App.css';

function App() {
  return (
    <InputField />
  );
}

function InputField() {
  const [text, setText] = useState('');

  const handleChange = (event) => {
    setText(event.target.value);
  };

  function getSortedExcellData() {
    const data = text;
    const splitData = data.trim().split('\n');
    for (let i = 0; i < splitData.length; i++) {
      if (splitData[i] === '') {
        splitData.splice(i, 1);
      }
    }

    let order;
    if (splitData.length === 4) {
      order = {
        ime: splitData[0],
        adresa: splitData[1],
        tel: splitData[2],
        cena: splitData[3],
      };
    }
    console.log(order);
    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet('Porudzbine');

    const headerRow = worksheet.addRow(['Ime i prezime', 'Adresa', 'Tel', 'Cena']);
    headerRow.eachCell((cell) => {
      cell.font = { bold: true };
      cell.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'FFFF00' },
      };
    });
    worksheet.addRow([order.ime, order.adresa, order.tel, order.cena]);

    worksheet.columns.forEach((column) => {
      let maxLength = 0;
      column.eachCell({ includeEmpty: true }, (cell) => {
        const length = cell.value ? cell.value.toString().length : 0;
        if (length > maxLength) {
          maxLength = length;
        }
      });
      column.width = maxLength < 10 ? 10 : maxLength + 2; // Set minimum width of 10
    });

    workbook.xlsx.writeBuffer().then((buffer) => {
      const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = 'products.xlsx';
      a.click();
      URL.revokeObjectURL(url);
    });
  }

  return (
    <div className="input-container">
      <label htmlFor="input-order">Insert Order Data</label>
      <textarea
        name="input-order"
        id="input-order"
        value={text}
        onChange={handleChange}
      />
      <button
        type="button"
        onClick={getSortedExcellData}
      >
        Get Data
      </button>
    </div>
  );
}

export default App;
