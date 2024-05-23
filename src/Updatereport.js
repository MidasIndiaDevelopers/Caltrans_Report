import React from 'react'
import * as XLSX from 'xlsx';
export const Updatereport = () => {
    const readFile = (file) => {
        const reader = new FileReader();
        reader.onload = (e) => {
            const data = new Uint8Array(e.target.result);
            const workbook = XLSX.read(data, { type: 'array' });
            const sheetName = workbook.SheetNames[0];
            const sheet = workbook.Sheets[sheetName];

            // Access cell value by name
            const cellValue = sheet['check1']; // Assuming 'Sales_Total' is a named cell
            console.log('Cell value:', cellValue);
        };
        reader.readAsArrayBuffer(file);
    };

    const handleFileUpload = (event) => {
        const file = event.target.files[0];
        if (file) {
            readFile(file);
        }
    };

    return (
        <div>
            <input type="file" onChange={handleFileUpload} />
        </div>
    );
}
