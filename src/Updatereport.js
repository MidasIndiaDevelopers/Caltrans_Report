import {
    DropList,
    Grid,
    Panel,
    Typography,
    VerifyUtil,
  } from "@midasit-dev/moaui";
  import { Radio, RadioGroup } from "@midasit-dev/moaui";
  import React, { useState, useEffect } from "react";
  import * as Buttons from "./Components/Buttons";
  import ExcelJS from "exceljs";
  import AlertDialogModal from "./AlertDialogModal";
  import { midasAPI } from "./Function/Common";
  import { enqueueSnackbar } from "notistack";
  import { ThetaBeta1 } from "./Function/ThetaBeta";
  import { ThetaBeta2 } from "./Function/ThetaBeta";
  import { TextField } from "@midasit-dev/moaui";
  import { saveAs } from "file-saver";
  
  export const Updatereport = () => {
    const [workbookData, setWorkbookData] = useState(null);
    const [sheetData, setSheetData] = useState([]);
    const [sheetName, setSheetName] = useState("");
    const [cast, setCast] = useState("inplace");
    const [sp, setSp] = useState("ca1");
    const [cvr, setCvr] = useState("ca2");
    const [value, setValue] = useState(1);
    const [SelectWorksheets, setWorksheet] = useState({});
    const [SelectWorksheets2, setWorksheet2] = useState({});
    const [SelectWorksheets3, setWorksheet3] = useState(null);
    const [lc, setLc] = useState({});
    const [item, setItem] = useState(new Map([["Select Load Combination", 1]]));
    const [check, setCheck] = useState(false);
    const [selectedName, setSelectedName] = useState("");
    const [matchedParts, setMatchedParts] = useState([]);
    let lcname;
    let mu_pos;
    let mu_neg;
    let mr_old_pos;
    let mr_new_pos;
    let mr_old_neg;
    let mr_new_neg;
  
  
  
    console.log(lcname);
    // let items = new Map([]);
  
    // For Summary File
    useEffect(() => {
      fetchLc();
    }, []);
    const fetchAndProcessExcelFile = async () => {
      try {
        const response = await fetch("/Summary_Caltrans.xlsx");
        if (!response.ok) {
          throw new Error("Network response was not ok.");
        }
        const arrayBuffer = await response.arrayBuffer();
  
        const workbook = new ExcelJS.Workbook();
        await workbook.xlsx.load(arrayBuffer);
  
        return workbook;
      } catch (error) {
        console.error("Error fetching or processing file:", error);
        throw error;
      }
    };
  
    console.log(selectedName);
  
    async function onChangeHandler(event) {
      setValue(event.target.value);
      console.log(event.target.value);
      // Find the name corresponding to the selected key
      for (let [name, key] of item.entries()) {
        if (key === event.target.value) {
          setSelectedName(name);
          lcname = name;
          console.log(lcname);
        }
      }
    }
  
  
    // const matchedParts = [];
    let worksheet;
    let worksheet2;
    let worksheet3;
    const handleFileUpload = async (event) => {
      const file = event.target.files[0];
      const reader = new FileReader();
      reader.onload = async (e) => {
        // await fetchLc();
  
        try {
          let buffer = e.target.result;
          let workbook = new ExcelJS.Workbook();
          await workbook.xlsx.load(buffer);
          let newMatchedParts = [];
          const summaryWorkbook = await fetchAndProcessExcelFile();
        const summarySheet = summaryWorkbook.getWorksheet('Sheet1');
  
        if (!summarySheet) {
          throw new Error("Sheet1 not found in Summary_Caltrans.xlsx");
        }
          const regex = /^([0-9]+)_([A-Z])$/;
  
          for (let key in workbook.worksheets) {
            let match = regex.exec(workbook.worksheets[key].name);
            if (match) {
              // match[1] contains the part that matches [0-9]+
              // match[2] contains the part that matches [A-Z]
              newMatchedParts.push({ numberPart: match[1], letterPart: match[2] });
  
              worksheet = workbook.worksheets[key];
              setWorksheet((prevState) => ({
                ...prevState,
                [key]: workbook.worksheets[key],
              }));
            }
            setMatchedParts(newMatchedParts); 
            console.log(matchedParts);
          }
  
          console.log(matchedParts);
          for (let key in workbook.worksheets) {
            const lcb = "StressAtLCB";
            if (workbook.worksheets[key].name === lcb) {
              worksheet2 = workbook.worksheets[key];
              setWorksheet2((prevstate) => ({
                ...prevstate,
                [key]: workbook.worksheets[key],
              }));
            }
          }
         
          console.log(worksheet);
          console.log(worksheet2);
         
  
          if (!worksheet) {
            throw new Error("No worksheets found in the uploaded file");
          } else {
            let cellvalue = worksheet.getRow(3).getCell(3).value;
            if (cellvalue != "AASHTO-LRFD2017") {
              alert("Incorrect file format");
            }
          }
  
          const indexToLetter = (index) => {
            let letter = "";
            while (index >= 0) {
              letter = String.fromCharCode((index % 26) + 65) + letter;
              index = Math.floor(index / 26) - 1;
            }
            return letter;
          };
  
          let startRowNumber1 = null;
          let endRowNumber1 = null;
          let startRowNumber2 = null;
          let endRowNumber2 = null;
  
          // Find start and end rows based on values in the first cell
          worksheet.eachRow((row, rowNumber) => {
            if (row.getCell(1).value === "$$strm1") {
              startRowNumber1 = rowNumber;
            } else if (row.getCell(1).value === "$$fpo") {
              endRowNumber1 = rowNumber;
            }
          });
          worksheet.eachRow((row, rowNumber) => {
            if (row.getCell(1).value === "$$strn1") {
              startRowNumber2 = rowNumber;
            } else if (row.getCell(1).value === "$$fpo_min") {
              endRowNumber2 = rowNumber;
            }
          });
  
          if (
            (startRowNumber1 === null || endRowNumber1 === null) &&
            (startRowNumber2 === null || endRowNumber2 === null)
          ) {
            throw new Error(
              "Could not find the start or end markers ($$strm1/$$strn1 or $$fpo/$$fpo_min)"
            );
          }
          // Process rows between startRowNumber and endRowNumber
          if (startRowNumber1 !== null && endRowNumber1 !== null) {
            for (
              let rowNumber = startRowNumber1 + 1;
              rowNumber < endRowNumber1;
              rowNumber++
            ) {
              let row = worksheet.getRow(rowNumber);
              row.eachCell({ includeEmpty: true }, (cell, colNumber) => {
                if (!cell.value) {
                  let colLetter = indexToLetter(colNumber - 1); // Adjusting index for 1-based column numbering
                  let address = colLetter + rowNumber;
                  row.getCell(colNumber).value = "";
                  row.getCell(colNumber)._address = address;
                }
              });
            }
          }
          if (startRowNumber2 !== null && endRowNumber2 !== null) {
            for (
              let rowNumber = startRowNumber2 + 1;
              rowNumber < endRowNumber2;
              rowNumber++
            ) {
              let row = worksheet.getRow(rowNumber);
              row.eachCell({ includeEmpty: true }, (cell, colNumber) => {
                if (colNumber !== 1) {
                  let colLetter = indexToLetter(colNumber - 1); // Adjusting index for 1-based column numbering
                  let address = colLetter + rowNumber;
                  row.getCell(colNumber).value = null; // Set cell value to null
                  row.getCell(colNumber)._address = address;
                }
              });
            }
          }
          const lastRowNumber = worksheet2.rowCount;
          const newRowNumber = lastRowNumber + 1;
          const newRow = worksheet2.getRow(newRowNumber);
  
          // Merge cells from column 1 to column 13
            worksheet2.mergeCells(newRowNumber, 1, newRowNumber + 1, 14);

            // Enter your content in the merged cell
            const mergedCell = worksheet2.getCell(newRowNumber, 1); // Get the top-left cell of the merged range
            mergedCell.value = "Tensile Stress unit in Prestress concrete at Service after loss: No tension case    (See CA-5.9.2.2b-1)";
            mergedCell.font = { bold: true, size: 12 }; // Increase the font size to 14

            // Center-align the content
            mergedCell.alignment = { vertical: 'middle' };
            newRow.commit(); // Commit the new row to the worksheet

            const targetSheet = workbook.addWorksheet('Summary');
            summarySheet.eachRow((row, rowNumber) => {
                const targetRow = targetSheet.getRow(rowNumber);
                row.eachCell((cell, colNumber) => {
                    targetRow.getCell(colNumber).value = cell.value;
                });
                targetRow.commit();
            });
            setWorkbookData(workbook);
            setSheetName(worksheet.name);
            console.log(workbook);
          for (let key in workbook.worksheets) {
              if (workbook.worksheets[key].name === "Summary") {
                worksheet3 = workbook.worksheets[key];
                setWorksheet3((prevstate) => ({
                  ...prevstate,
                  [key]: workbook.worksheets[key],
                }));
              }
            }
            console.log(worksheet3);
        } catch (error) {
          console.error("Error reading file:", error);
          alert(
            "Error reading file. Please make sure the file is a valid Excel file."
          );
        }
      };
  
      reader.readAsArrayBuffer(file);
    };
    console.log(matchedParts);
    const [ag, setAg] = useState("");
    const [sg, setSg] = useState("");
  
    const handleAgChange = (event) => {
      setAg(event.target.value);
    };
  
    const handleSgChange = (event) => {
      setSg(event.target.value);
    };
    console.log(ag);
    console.log(sg);
    async function updatedata(wkey, worksheet) {
      if (!workbookData) return;
      if (!worksheet) {
        throw new Error("No worksheets found in the uploaded file");
      }
      
      let rows = worksheet._rows;
      let mn;
      let mn_neg;
      let phi;
      let mr;
      let mr_neg;
      let dv;
      let dv_min;
      let data = {};
      let Av;
      let Avm;
      let Mmax;
      let Mmin;
      let Ag;
      let St;
      let Sb;
      let Nmax;
      let Nmin;
      let E;
      let fc;
      let Vu1;
      let Vu2;
      let beta1;
      let Vc;
      let Vc_min;
      let Vc1;
      let dc;
      let storedValues = {};
      let pi = null;
      let pi_min = null;
      let s_max;
      let s_min;
      let $$alpha;
      let $$alpha_min;
      let Vp;
      let Vp_min;
      let type;
      let K;
      let a;
      let Vu_max;
      let Vu_min;
      let fy;
  
      for (let key1 in rows) {
        if (
          rows[key1]._cells[0] != undefined &&
          rows[key1]._cells[0]._value.model.value == "$$type"
        ) {
          let cell17 = rows[key1]._cells[17];
          let add17 = cell17._address;
          let cell17Value = cell17.value !== undefined ? cell17.value : null;
          if (cell17Value == "Composite") {
            type = "Composite";
          } else {
            type = "Box";
          }
        }
  
        if (
          rows[key1]._cells[0] != undefined &&
          rows[key1]._cells[0]._value.model.value == "$$Mn"
        ) {
          let location = rows[key1]._cells[19]._value.model.address;
          let value = rows[key1]._cells[19]._value.model.value;
          data = { ...data, [location]: value };
          mn = value;
        }
  
        // to get Phi row
        if (
          rows[key1]._cells[0] != undefined &&
          rows[key1]._cells[0]._value.model.value == "$$Phi"
        ) {
          let location = rows[key1]._cells[5]._value.model.address;
          if (cast === "inplace") {
            data = { ...data, [location]: 0.95 };
            phi = 0.95;
            let equ = rows[key1]._cells[22]._value.model.address;
            // Retrieve the existing value from cell[22]
            let existingValue = rows[key1]._cells[22]._value.model.value;
            // Concatenate the existing value with the new string
            let concatenatedValue = '0.005 ≤εt  ' + ' (See CA- 5.5.4.2)';
            data = { ...data, [equ]: concatenatedValue };
          } else {
            data = { ...data, [location]: 1 };
            phi = 1;
          }
          console.log(phi);
        }
        if (
          rows[key1]._cells[0] != undefined &&
          rows[key1]._cells[0]._value.model.value == "$$Mr"
        ) {
          let location = rows[key1]._cells[5]._value.model.address;
  
          let mu = rows[key1]._cells[17]._value.model.value;
          mu_pos = mu;
          mr_old_pos = rows[key1]._cells[5]._value.model.value;
          mr = Math.round(Number(mn) * Number(phi) * 100) / 100;
          mr_new_pos = mr;
          data = { ...data, [location]: mr };
  
          // location of oK
          if (mr < Number(mu)) {
            let location1 = rows[key1]._cells[29]._value.model.address;
            let location2 = rows[key1]._cells[13]._value.model.address;
            data = { ...data, [location1]: "NG" };
            data = { ...data, [location2]: "<" };
          }
          else {
              let location1 = rows[key1]._cells[29]._value.model.address;
              let location2 = rows[key1]._cells[13]._value.model.address;
              data = { ...data, [location1]: "OK" };
              data = { ...data, [location2]: "≥" };
          }
        }
        if (
          rows[key1]._cells[0] != undefined &&
          rows[key1]._cells[0]._value.model.value == "$$Mn_min"
        ) {
          let location = rows[key1]._cells[19]._value.model.address;
          let value = rows[key1]._cells[19]._value.model.value;
          data = { ...data, [location]: value };
          mn_neg = value;
        }
  
        // to get Phi row
        if (
          rows[key1]._cells[0] != undefined &&
          rows[key1]._cells[0]._value.model.value == "$$Phi_min"
        ) {
          let location = rows[key1]._cells[5]._value.model.address;
          if (cast === "inplace") {
            data = { ...data, [location]: 0.95 };
            phi = 0.95;
            let equ = rows[key1]._cells[25]._value.model.address;
            data = { ...data, [equ] : '(See CA- 5.5.4.2)'}
          } else {
            data = { ...data, [location]: 1 };
            phi = 1;
          }
          console.log(phi);
        }
        if (
          rows[key1]._cells[0] != undefined &&
          rows[key1]._cells[0]._value.model.value == "$$Mr_min"
        ) {
          let location = rows[key1]._cells[5]._value.model.address;
  
          let mu = rows[key1]._cells[17]._value.model.value;
          mu_neg =mu;
          mr_neg = Math.round(Number(mn_neg) * Number(phi) * 100) / 100;
          data = { ...data, [location]: mr };
  
          // location of oK
          if (mr_neg < Number(mu)) {
            let location1 = rows[key1]._cells[29]._value.model.address;
            let location2 = rows[key1]._cells[13]._value.model.address;
            data = { ...data, [location1]: "NG" };
            data = { ...data, [location2]: "<" };
          }
          else {
              let location1 = rows[key1]._cells[29]._value.model.address;
              let location2 = rows[key1]._cells[13]._value.model.address;
              data = { ...data, [location1]: "OK" };
              data = { ...data, [location2]: "≥" };
          }
        }
        if (
          rows[key1]._cells[0] != undefined &&
          rows[key1]._cells[0]._value.model.value == "$$dv"
        ) {
          dv = rows[key1]._cells[4]._value.model.value;
        }
        if (
          rows[key1]._cells[0] != undefined &&
          rows[key1]._cells[0]._value.model.value == "$$dv_min"
        ) {
          dv_min = rows[key1]._cells[4]._value.model.value;
        }
        if (
          rows[key1]._cells[0] != undefined &&
          rows[key1]._cells[0]._value.model.value == "$$sm"
        ) {
          if (sp === "ca1") {
            let add1 = rows[key1]._cells[6]._value.model.address;
            data = { ...data, [add1]: "Min[0.8dv, 18.0(in.)]" };
            let add2 = rows[key1]._cells[13]._value.model.address;
            // let val=rows[key1]._cells[13]._value.model.value;
            if (0.8 * dv >= 18) {
              data = { ...data, [add2]: 18 };
            } else {
              data = { ...data, [add2]: 0.8 * dv };
            }
          }
          let add27 = rows[key1]._cells[27]._value.model.address;
          data = { ...data,[add27] : '(See CA-5.7.2.6-1)'}
        }
        if (
          rows[key1]._cells[0] != undefined &&
          rows[key1]._cells[0]._value.model.value == "$$s"
        ) {
          let cell8 = rows[key1]._cells[8];
          s_max = rows[key1]._cells[8].value;
          console.log(s_max);
        }
        if (
          rows[key1]._cells[0] != undefined &&
          rows[key1]._cells[0]._value.model.value == "$$s_min"
        ) {
          let cell6 = rows[key1]._cells[6];
          s_min = rows[key1]._cells[6].value;
          console.log(s_min);
        }
        if (
          rows[key1]._cells[0] != undefined &&
          rows[key1]._cells[0]._value.model.value == "$$SX"
        ) {
          K = rows[key1]._cells[15].value;
        }
        if (
          rows[key1]._cells[0] != undefined &&
          rows[key1]._cells[0]._value.model.value == "$$dc"
        ) {
          if (cvr === "ca2") {
            for (let col = 0; col < rows[key1]._cells.length; col++) {
              if (rows[key1]._cells[col] != undefined) {
                let cellValue = rows[key1]._cells[col]._value.model.value;
                let cellAddress = rows[key1]._cells[col]._value.model.address;
                storedValues[cellAddress] = cellValue;
              }
            }
            let add1 = rows[key1]._cells[8]._value.model.address;
            let add2 = rows[key1]._cells[11]._value.model.address;
            let val2 = rows[key1]._cells[11]._value.model.value;
            let val2_new;
            let val3 = rows[key1]._cells[21]._value.model.value;
            let add4 = rows[key1]._cells[29]._value.model.address;
            let add5 = rows[key1]._cells[17]._value.model.address;
            let column15Address;
            let column15Value;
            let column15Value_new;
            let column9Address;
            let column9Value;
            let column9Value_new;
  
            // Move to the next row and check for $$B and the next $$dc
            let nextKey1 = parseInt(key1, 10) + 1;
            while (rows[nextKey1]) {
              if (rows[nextKey1]._cells[0] != undefined) {
                if (rows[nextKey1]._cells[0]._value.model.value == "$$B") {
                  // Store the value and address of cell in column 13 for $$B row
                  column15Address =
                    rows[nextKey1]._cells[15]._value.model.address;
                  column15Value = rows[nextKey1]._cells[15]._value.model.value;
                  column15Value_new =
                    1 +
                    (1 * 2.5) / (1 / ((column15Value - 1) / 1.8) + 1.26 - 1.75);
                  column15Value_new = Math.round(column15Value_new);
                  console.log(column15Value_new);
                  storedValues[column15Address] = column15Value;
                  data = { ...data, [column15Address]: column15Value_new };
                }
                if (rows[nextKey1]._cells[0]._value.model.value == "$$d-c") {
                  // Store the value and address of cell in column 9 for $$dc row
                  column9Value = rows[nextKey1]._cells[9]._value.model.value;
                  column9Address = rows[nextKey1]._cells[9]._value.model.address;
                  column9Value_new = column9Value - column9Value + 2.5;
                  console.log(column9Value_new);
                  storedValues[column9Address] = column9Value;
                  data = { ...data, [column9Address]: column9Value_new };
                  let add13 = rows[nextKey1]._cells[13]._value.model.address;
                  data = { ...data,[add13] : '(in)      (See CA-5.6.7-1)'};
                  break;
                }
              }
              nextKey1++;
            }
            console.log(storedValues);
            val2_new = ((val2 + 3.6) * column15Value) / column15Value_new - 5;
  
            if (val2_new < 0) {
              val2_new = 0.0;
            }
  
            if (val2_new < val3) {
              data = { ...data, [add4]: "NG" };
              data = { ...data, [add5]: "<" };
            }
  
            data = { ...data, [add1]: "2*2.5" };
            data = { ...data, [add2]: val2_new };
          }
        }
        if (
          rows[key1]._cells[0] != undefined &&
          rows[key1]._cells[0]._value.model.value == "$$a"
        ) {
          let cell8 = rows[key1]._cells[8];
          a = cell8.value;
        }
        if (
          rows[key1]._cells[0] != undefined &&
          rows[key1]._cells[0]._value.model.value == "$$a_min"
        ) {
          let cell8 = rows[key1]._cells[8];
          a = cell8.value;
        }
        console.log(a);
        if (
          rows[key1]._cells[0] != undefined &&
          rows[key1]._cells[0]._value.model.value == "$$fy"
        ) {
          fy = rows[key1]._cells[6].value;
        }
        console.log(fy);
  
        if (
          rows[key1]._cells[0] != undefined &&
          rows[key1]._cells[0]._value.model.value == "$$Avm"
        ) {
          if (type == "Composite") {
            Avm = rows[key1]._cells[12]._value.model.value;
          } else {
            Avm = rows[key1]._cells[17]._value.model.value;
          }
        }
        if (
          rows[key1]._cells[0] != undefined &&
          rows[key1]._cells[0]._value.model.value == "$$Av"
        ) {
          Av = rows[key1]._cells[5]._value.model.value;
        }
        if (
          rows[key1]._cells[0] != undefined &&
          rows[key1]._cells[0]._value.model.value == "$$Mmax"
        ) {
          Mmax = rows[key1]._cells[15]._value.model.value;
        }
        if (
          rows[key1]._cells[0] != undefined &&
          rows[key1]._cells[0]._value.model.value == "$$Mmin"
        ) {
          Mmin = rows[key1]._cells[15]._value.model.value;
        }
        if (
          rows[key1]._cells[0] != undefined &&
          rows[key1]._cells[0]._value.model.value == "$$Ag"
        ) {
          Ag = rows[key1]._cells[27]._value.model.value;
        }
        if (
          rows[key1]._cells[0] != undefined &&
          rows[key1]._cells[0]._value.model.value == "$$St"
        ) {
          St = rows[key1]._cells[27]._value.model.value;
        }
        if (
          rows[key1]._cells[0] != undefined &&
          rows[key1]._cells[0]._value.model.value == "$$Sb"
        ) {
          Sb = rows[key1]._cells[27]._value.model.value;
        }
        if (
          rows[key1]._cells[0] != undefined &&
          rows[key1]._cells[0]._value.model.value == "$$Nmax"
        ) {
          Nmax = rows[key1]._cells[15]._value.model.value;
        }
        if (
          rows[key1]._cells[0] != undefined &&
          rows[key1]._cells[0]._value.model.value == "$$Nmin"
        ) {
          Nmin = rows[key1]._cells[15]._value.model.value;
        }
        if (
          rows[key1]._cells[0] != undefined &&
          rows[key1]._cells[0]._value.model.value == "$$E"
        ) {
          E = rows[key1]._cells[12]._value.model.value;
          fc = rows[key1]._cells[5]._value.model.value;
        }
        if (
          rows[key1]._cells[0] != undefined &&
          rows[key1]._cells[0]._value.model.value == "$$Vu1"
        ) {
          if (type == "Composite") {
            Vu1 = rows[key1]._cells[10]._value.model.value;
          } else {
            Vu1 = rows[key1]._cells[13]._value.model.value;
          }
        }
        if (
          rows[key1]._cells[0] != undefined &&
          rows[key1]._cells[0]._value.model.value == "$$Vu2"
        ) {
          if (type == "Composite") {
            Vu2 = rows[key1]._cells[10]._value.model.value;
          } else {
            Vu2 = rows[key1]._cells[13]._value.model.value;
          }
        }
        if (
          rows[key1]._cells[0] != undefined &&
          rows[key1]._cells[0]._value.model.value == "$$Vc"
        ) {
          Vc = rows[key1]._cells[9]._value.model.value;
        }
        if (
          rows[key1]._cells[0] != undefined &&
          rows[key1]._cells[0]._value.model.value == "$$alpha"
        ) {
          let cell8 = rows[key1]._cells[8];
          $$alpha = rows[key1]._cells[8].value;
        }
        if (
          rows[key1]._cells[0] != undefined &&
          rows[key1]._cells[0]._value.model.value == "$$alpha_min"
        ) {
          let cell8 = rows[key1]._cells[8];
          $$alpha_min = rows[key1]._cells[8].value;
        }
        if (
          rows[key1]._cells[0] != undefined &&
          rows[key1]._cells[0]._value.model.value == "$$Vp"
        ) {
          let cell19 = rows[key1]._cells[19];
          Vp = rows[key1]._cells[19].value;
        }
        if (
          rows[key1]._cells[0] != undefined &&
          rows[key1]._cells[0]._value.model.value == "$$Vp_min"
        ) {
          let cell19 = rows[key1]._cells[19];
          Vp_min = rows[key1]._cells[19].value;
        }
      }
      console.log(Ag);
      console.log(Sb);
      console.log(St);
      console.log(E);
      console.log(fc);
      console.log(Avm);
      console.log(Av);
      console.log(Mmax);
      console.log(Mmin);
      console.log(Nmax);
      console.log(Nmin);
      console.log(Vu1);
      console.log(Vu2);
      console.log(Vc);
      console.log($$alpha);
      console.log(Vp);
      let Ecm =
        ((-1 * Number(Mmax)) / Number(St) + Number(Nmax) / Number(Ag)) /
        Number(E);
      let Ecn =
        ((-1 * Number(Mmin)) / Number(St) + Number(Nmin) / Number(Ag)) /
        Number(E);
      let Etm =
        ((-1 * Number(Mmax)) / Number(Sb) + Number(Nmax) / Number(Ag)) /
        Number(E);
      let Etn =
        ((-1 * Number(Mmin)) / Number(Sb) + Number(Nmin) / Number(Ag)) /
        Number(E);
      // console.log(Vu1,Vu2,fc)
      let a1 = Number(Vu1) / Number(fc);
      let a2 = Number(Vu2) / Number(fc);
      let Exm = (Ecm + Etm) / 2;
      let Exn = (Ecn + Etn) / 2;
      // console.log(a1,a2, Exm * 1000, Exn * 1000)
      let value1 = ThetaBeta1(a1, Exm * 1000);
      let value2 = ThetaBeta1(a2, Exn * 1000);
      // console.log(value1,value2);
      let theta1 = value1[0];
      let theta2 = value2[0];
      beta1 = value1[1];
      let beta2 = value2[1];
      Vc1 = Vc / beta1;
      console.log(theta1, theta2, beta1, beta2);
      let startBlanking = false;
      let beta;
      let beta_min;
      let half_finalResult;
      let finalResult;
      let Vn;
      let Vn_min;
      let Vs;
      let Vs_min;
      let theta_new;
      let theta_new_min;
      let beta_new;
      let beta_new_min;
      function indexToLetter(index) {
        let letter = "";
        while (index >= 0) {
          letter = String.fromCharCode((index % 26) + 65) + letter;
          index = Math.floor(index / 26) - 1;
        }
        return letter;
      }
      let initialValue13 = null;
      let newValue13 = null;
      for (let key1 in rows) {
          if (rows[key1]._cells[0] != undefined && rows[key1]._cells[0]._value.model.value == "Design Condition") {
              // let add1 = rows[key1]._cells[2]._value.model.address;
              // let add2 = rows[key1]._cells[8]._value.model.address;
              // data = { ...data, [add1]: "Calculation for β and θ" };
              // data = { ...data, [add2]: "" };
              let add1 = rows[key1]._cells[0]._value.model.address;
              data = { ...data ,[add1] : 'Design Condition for Caltrans Amendment As per AASHTO LRFD Bridge Design'}
            }
  
        if (
          rows[key1]._cells[0] != undefined &&
          rows[key1]._cells[0]._value.model.value == "$$strm"
        ) {
          let add1 = rows[key1]._cells[2]._value.model.address;
          let add2 = rows[key1]._cells[8]._value.model.address;
          data = { ...data, [add1]: "Calculation for β and θ  (See CA - 5.7.3.4)" };
          data = { ...data, [add2]: "" };
          // let add12 = rows[key1]._cells[12]._value.model.address;
          // data = { ...data, [add12] : '(See CA - 5.7.3.4)'}
        }
        let cell;
        let add15value = dv < sg ? dv : sg;
        let sxe = (add15value * 1.38) / (ag + 0.63);
        if (
          rows[key1]._cells[0] != undefined &&
          rows[key1]._cells[0]._value.model.value == "$$strm1"
        ) {
          console.log(rows[key1]);
  
          for (let i = 1; i <= 50; i++) {
            // Check if the cell is present
            if (!rows[key1]._cells[i]) {
              // Create a dummy cell if it's missing
              rows[key1]._cells[i] = {
                _value: {
                  model: {
                    value: "dummy",
                    address: indexToLetter(i) + (parseInt(key1) + 1), // Assign a dummy value
                    // Add other necessary properties if required
                  },
                },
  
                _address: indexToLetter(i) + (parseInt(key1) + 1),
              };
            } else {
              // If the cell is present, clear its value
              cell = rows[key1]._cells[i];
              if (cell && cell.model) {
                cell.model = {};
                cell.model.value = " "; // Clear the value
              }
            }
          }
  
          // Run another while loop to make corresponding column cells empty until $$fpo is found
          let nextKey = parseInt(key1) + 1;
          while (rows[nextKey] !== undefined) {
            console.log(rows[nextKey]);
            if (
              rows[nextKey]._cells[0] != undefined &&
              rows[nextKey]._cells[0]._value.model.value == "$$fpo"
            ) {
              break; // Stop the loop when $$fpo is found
            }
            // Blank all cells in the current row of the while loop
            for (let i = 1; i <= 50; i++) {
              // Check if the cell is present
              if (!rows[nextKey]._cells[i]) {
                // Create a dummy cell if it's missing
                rows[nextKey]._cells[i] = {
                  _value: {
                    model: {
                      value: "dummy", // Assign a dummy value
                      address: indexToLetter(i) + (nextKey + 1), // Add other necessary properties if required
                    },
                  },
                  _address: indexToLetter(i) + (nextKey + 1),
                };
              } else {
                // If the cell is present, clear its value
                cell = rows[nextKey]._cells[i];
                if (cell && cell.model) {
                  cell.model = {};
                  cell.model.value = " "; // Clear the value
                }
              }
            }
  
            nextKey++;
          }
        }
        if (
          rows[key1]._cells[0] != undefined &&
          rows[key1]._cells[0]._value.model.value == "$$strn1"
        ) {
          console.log(rows[key1]);
  
          for (let i = 1; i <= 50; i++) {
            // Check if the cell is present
            if (!rows[key1]._cells[i]) {
              // Create a dummy cell if it's missing
              rows[key1]._cells[i] = {
                _value: {
                  model: {
                    value: "dummy",
                    address: indexToLetter(i) + (parseInt(key1) + 1), // Assign a dummy value
                    // Add other necessary properties if required
                  },
                },
  
                _address: indexToLetter(i) + (parseInt(key1) + 1),
              };
            } else {
              // If the cell is present, clear its value
              cell = rows[key1]._cells[i];
              if (cell && cell.model) {
                cell.model = {};
                cell.model.value = " "; // Clear the value
              }
            }
          }
  
          // Run another while loop to make corresponding column cells empty until $$fpo is found
          let nextKey = parseInt(key1) + 1;
          while (rows[nextKey] !== undefined) {
            console.log(rows[nextKey]);
            if (
              rows[nextKey]._cells[0] != undefined &&
              rows[nextKey]._cells[0]._value.model.value == "$$fpo_min"
            ) {
              break; // Stop the loop when $$fpo is found
            }
            // Blank all cells in the current row of the while loop
            for (let i = 1; i <= 50; i++) {
              // Check if the cell is present
              if (!rows[nextKey]._cells[i]) {
                // Create a dummy cell if it's missing
                rows[nextKey]._cells[i] = {
                  _value: {
                    model: {
                      value: "dummy", // Assign a dummy value
                      address: indexToLetter(i) + (nextKey + 1), // Add other necessary properties if required
                    },
                  },
                  _address: indexToLetter(i) + (nextKey + 1),
                };
              } else {
                // If the cell is present, clear its value
                cell = rows[nextKey]._cells[i];
                if (cell && cell.model) {
                  cell.model = {};
                  cell.model.value = " "; // Clear the value
                }
              }
            }
  
            nextKey++;
          }
        }
        if (
          rows[key1]._cells[0] != undefined &&
          rows[key1]._cells[0]._value.model.value == "$$pi"
        ) {
          if (
            rows[key1]._cells[15] &&
            rows[key1]._cells[15].value !== undefined
          ) {
            pi = rows[key1]._cells[15].value;
          } else {
            console.error(
              "Error: Unable to retrieve value for rows[key1]._cells[15]"
            );
          }
        }
        console.log("Value of pi:", pi);
        if (
          rows[key1]._cells[0] != undefined &&
          rows[key1]._cells[0]._value.model.value == "$$pi_min"
        ) {
          if (
            rows[key1]._cells[15] &&
            rows[key1]._cells[15].value !== undefined
          ) {
            pi_min = rows[key1]._cells[15].value;
          } else {
            console.error(
              "Error: Unable to retrieve value for rows[key1]._cells[15]"
            );
          }
        }
        console.log("Value of pi:", pi);
  
        if (
          rows[key1]._cells[0] != undefined &&
          rows[key1]._cells[0]._value.model.value == "$$A"
        ) {
          console.log(rows[key1]._cells);
          let cell = rows[key1]._cells[4];
          if (cell && cell._address) {
            let add11 = cell._address;
            data = { ...data, [add11]: "A" };
          } else {
            console.error(
              "Error: Unable to determine address for rows[key1]._cells[4]"
            );
          }
          let cell2 = rows[key1]._cells[6];
          if (cell2 && cell2._address) {
            let add12 = cell2._address;
            data = { ...data, [add12]: "Aₘᵢₙ" };
          } else {
            console.error(
              "Error: Unable to determine address for rows[key1]._cells[4]"
            );
          }
          let cell3 = rows[key1]._cells[5];
          let comparisonSymbol = Av >= Avm ? "≥" : "<";
          if (cell3 && cell3._address) {
            let add13 = cell3._address;
            data = { ...data, [add13]: comparisonSymbol };
          }
        }
        if (
          rows[key1]._cells[0] != undefined &&
          rows[key1]._cells[0]._value.model.value == "$$e"
        ) {
          console.log(rows[key1]._cells);
          let cell = rows[key1]._cells[4];
          if (cell && cell._address) {
            let add11 = cell._address;
            data = { ...data, [add11]: "εx" };
          } else {
            // Handle the case where _address is undefined or not available
            console.error(
              "Error: Unable to determine address for rows[key1]._cells[4]"
            );
            // Handle this error scenario appropriately
          }
          let cell2 = rows[key1]._cells[5];
  
          // Check if cell and its properties are defined
          if (cell2 && cell2._address) {
            let add12 = cell2._address;
            data = { ...data, [add12]: "=" };
          } else {
            // Handle the case where _address is undefined or not available
            console.error(
              "Error: Unable to determine address for rows[key1]._cells[4]"
            );
            // Handle this error scenario appropriately
          }
          let cell3 = rows[key1]._cells[6];
          if (cell3 && cell3._address) {
            let add13 = cell3._address;
            let cell3Value = Av >= Avm ? Exm : Etm;
            data = { ...data, [add13]: cell3Value };
          }
        }
        if (
          rows[key1]._cells[0] != undefined &&
          rows[key1]._cells[0]._value.model.value == "$$sx"
        ) {
          rows[key1]._cells = rows[key1]._cells.map((cell) =>
            cell === "" ? undefined : cell
          );
          console.log(rows[key1]._cells);
          let cell = rows[key1]._cells[4];
          if (cell && cell._address) {
            let add11 = cell._address;
            data = { ...data, [add11]: "sx" };
          } else {
            console.error(
              "Error: Unable to determine address for rows[key1]._cells[4]"
            );
          }
  
          let cell2 = rows[key1]._cells[5];
          if (cell2 && cell2._address) {
            let add12 = cell2._address;
            data = { ...data, [add12]: "=" };
          } else {
            console.error(
              "Error: Unable to determine address for rows[key1]._cells[5]"
            );
          }
  
          let startCol = 7;
          let endCol = 17;
          let rowNumber = rows[key1]._cells[startCol].row;
          try {
            let mergeRange = worksheet.getCell(
              `${worksheet.getColumn(startCol).letter}${rowNumber}:${
                worksheet.getColumn(endCol).letter
              }${rowNumber}`
            );
            if (!mergeRange.isMerged) {
              worksheet.mergeCells(rowNumber, startCol, rowNumber, endCol);
            }
          } catch (error) {
            console.error("Error merging cells: ", error);
          }
  
          let cell3 = worksheet.getCell(rowNumber, startCol);
          if (cell3 && cell3._address) {
            let add13 = cell3._address;
  
            if (Av < Avm) {
              data = {
                ...data,
                [add13]: `Min| dv, maximum distance between the longitudinal r/f |`,
              };
              let cell4 = rows[key1]._cells[18];
              if (cell4 && cell4._address) {
                let add14 = cell4._address;
                data = { ...data, [add14]: "=" };
              } else {
                console.error(
                  "Error: Unable to determine address for rows[key1]._cells[24]"
                );
              }
  
              let cell5 = rows[key1]._cells[19];
              if (cell5 && cell5._address) {
                let add15 = cell5._address;
                let add15value = dv < sg ? dv : sg;
                data = { ...data, [add15]: add15value };
              } else {
                console.error(
                  "Error: Unable to determine address for rows[key1]._cells[25]"
                );
              }
            } else {
              data = { ...data, [add13]: `Not Required` };
            }
          } else {
            console.error(
              "Error: Unable to determine address for rows[key1]._cells[24]"
            );
          }
        }
  
        if (
          rows[key1]._cells[0] != undefined &&
          rows[key1]._cells[0]._value.model.value == "$$sxe"
        ) {
          let cell3 = rows[key1]._cells[4];
          if (cell3 && cell3._address) {
            let add13 = cell3._address;
            // let sxe = ((add15value*1.38)/(ag +0.63))
            data = { ...data, [add13]: "sxe" };
          }
          let cell4 = rows[key1]._cells[5];
          if (cell4 && cell4._address) {
            let add14 = cell4._address;
            // let sxe = ((add15value*1.38)/(ag +0.63))
            data = { ...data, [add14]: "=" };
          }
          let mergeStartCol = 7;
          let mergeEndCol = 16;
          let mergeRowNumber = rows[key1]._cells[mergeStartCol].row;
  
          // Check if the range is already merged
          try {
            let mergeRange = worksheet.getCell(
              `${worksheet.getColumn(mergeStartCol).letter}${mergeRowNumber}:${
                worksheet.getColumn(mergeEndCol).letter
              }${mergeRowNumber}`
            );
            if (!mergeRange.isMerged) {
              worksheet.mergeCells(
                mergeRowNumber,
                mergeStartCol,
                mergeRowNumber,
                mergeEndCol
              );
            }
          } catch (error) {
            console.error("Error merging cells: ", error);
          }
  
          // After merging, the cell5 should refer to the merged cell
          let cell5 = worksheet.getCell(mergeRowNumber, mergeStartCol);
  
          if (cell5 && cell5._address) {
            let add15 = cell5._address;
  
            if (Av < Avm) {
              data = { ...data, [add15]: sxe };
            } else {
              data = { ...data, [add15]: "Not required" };
            }
          } else {
            console.error(
              "Error: Unable to determine address for rows[key1]._cells[6]"
            );
          }
        }
  
        if (
          rows[key1]._cells[0] != undefined &&
          rows[key1]._cells[0]._value.model.value == "$$b"
        ) {
          let cell5 = rows[key1]._cells[4];
          let add5 = cell5._address;
          data = { ...data, [add5]: "β" };
          let cell6 = rows[key1]._cells[5];
          let add6 = cell6._address;
          data = { ...data, [add6]: "=" };
          let cell7 = rows[key1]._cells[6];
          let add7 = cell7._address;
          if (Av < Avm) {
            let b_value = ThetaBeta2(sxe, Etm * 1000);
            let beta = b_value[0];
            console.log(beta);
            data = { ...data, [add7]: beta };
            beta_new = beta;
          } else {
            data = { ...data, [add7]: parseFloat(beta1.toFixed(2)) };
            beta_new = parseFloat(beta1.toFixed(2));
          }
          let cell15 = rows[key1]._cells[15];
          let add15 = cell15._address;
          data = { ...data,[add15] : '(See CA-5.7.3.4)'}
        }
  
        if (
          rows[key1]._cells[0] != undefined &&
          rows[key1]._cells[0]._value.model.value == "$$theta_max"
        ) {
          let cell5 = rows[key1]._cells[4];
          let add5 = cell5._address;
          data = { ...data, [add5]: "θ" };
          let cell6 = rows[key1]._cells[5];
          let add6 = cell6._address;
          data = { ...data, [add6]: "=" };
          let cell7 = rows[key1]._cells[6];
          let add7 = cell7._address;
          if (Av < Avm) {
            let theta_value = ThetaBeta2(sxe, Etm * 1000);
            let theta = theta_value[1];
            console.log(theta);
            data = { ...data, [add7]: theta };
            theta_new = theta;
          } else {
            data = { ...data, [add7]: parseFloat(theta1.toFixed(2)) };
            theta_new = parseFloat(theta1.toFixed(2));
            let cell15 = rows[key1]._cells[15];
            let add15 = cell15._address;
            data = { ...data,[add15] : '(See CA-5.7.3.4)'}
          }
        }
  
        if (
          rows[key1]._cells[0] != undefined &&
          rows[key1]._cells[0]._value.model.value == "$$beta"
        ) {
          let cell9 = rows[key1]._cells[9];
          if (cell9 && cell9._address) {
            beta = cell9._value.model.value;
          } else {
            // Handle the case where _address is undefined or not available
            console.error(
              "Error: Unable to determine address for rows[key1]._cells[9]"
            );
            // Handle this error scenario appropriately
          }
        }
        if (
          rows[key1]._cells[0] != undefined &&
          rows[key1]._cells[0]._value.model.value == "$$beta_min"
        ) {
          let cell9 = rows[key1]._cells[9];
          if (cell9 && cell9._address) {
            beta_min = cell9._value.model.value;
          } else {
            // Handle the case where _address is undefined or not available
            console.error(
              "Error: Unable to determine address for rows[key1]._cells[9]"
            );
            // Handle this error scenario appropriately
          }
        }
        if (
          rows[key1]._cells[0] != undefined &&
          rows[key1]._cells[0]._value.model.value == "$$vc"
        ) {
          let cell13 = rows[key1]._cells[13];
          let cell9 = rows[key1]._cells[9];
          let cell5 = rows[key1]._cells[5];
          let add5 = cell5._address;
          let add9 = cell9._address;
          data = { ...data, [add5]: "0.0316" };
          console.log(rows[key1]._cells[9].value);
          console.log(rows[key1]);
          if (cell13 && cell13._address) {
            if (type == "Composite") {
              let add13 = cell13._address;
              // Store the initial value globally
              initialValue13 = cell13._value.model.value;
              // Retrieve the value from cell13, divide it by beta, and multiply by beta1
              let value13 = initialValue13;
              let result = (value13 / beta) * beta1;
              // Store the result back in cell13
              cell13._value.model.value = result;
              // Store the new value globally
              newValue13 = result; //new Vc value
              data = { ...data, [add9]: "β √f'c bvdv" };
              data = { ...data, [add13]: parseFloat(result.toFixed(2)) };
              Vc = newValue13;
            } else {
              let add13 = cell13._address;
              // Store the initial value globally
              initialValue13 = cell13._value.model.value;
              // Retrieve the value from cell13, divide it by beta, and multiply by beta1
              let value13 = initialValue13;
              let result = (value13 / K) * beta1;
              // Store the result back in cell13
              cell13._value.model.value = result;
              // Store the new value globally
              newValue13 = result; //new Vc value
              data = { ...data, [add9]: "β √f'c bvdv" };
              data = { ...data, [add13]: parseFloat(result.toFixed(2)) };
              Vc = newValue13;
            }
          } else {
            // Handle the case where _address is undefined or not available
            console.error(
              "Error: Unable to determine address for rows[key1]._cells[13]"
            );
            // Handle this error scenario appropriately
          }
        }
        console.log("Initial Value:", initialValue13);
        console.log("New Value:", newValue13);
        if (
          rows[key1]._cells[0] != undefined &&
          rows[key1]._cells[0]._value.model.value == "$$vc_min"
        ) {
          let vc_o;
          let cell13 = rows[key1]._cells[13];
          vc_o = cell13.value;
          let vc_n;
          let cell9 = rows[key1]._cells[9];
          let cell5 = rows[key1]._cells[5];
          let add5 = cell5._address;
          let add9 = cell9._address;
          data = { ...data, [add5]: "0.0316" };
          console.log(rows[key1]._cells[9].value);
          console.log(rows[key1]);
          if (cell13 && cell13._address) {
            if (type == "Composite") {
              let add13 = cell13._address;
              // Store the initial value globally
              vc_o = cell13._value.model.value;
              // Retrieve the value from cell13, divide it by beta, and multiply by beta1
              let result = (vc_o / beta_min) * beta2;
              // Store the result back in cell13
              cell13._value.model.value = result;
              // Store the new value globally
              vc_n = result; //new Vc value
              data = { ...data, [add9]: "β √f'c bvdv" };
              data = { ...data, [add13]: parseFloat(vc_n.toFixed(2)) };
              Vc_min = newValue13;
            } else {
              let add13 = cell13._address;
              // Store the initial value globally
              vc_o = cell13._value.model.value;
              // Retrieve the value from cell13, divide it by beta, and multiply by beta1
              let value13 = vc_o;
              let result = (value13 / K) * beta2;
              // Store the result back in cell13
              cell13._value.model.value = result;
              // Store the new value globally
              vc_n = result; //new Vc value
              data = { ...data, [add9]: "β √f'c bvdv" };
              data = { ...data, [add13]: parseFloat(vc_n.toFixed(2)) };
              Vc_min = vc_n;
            }
          } else {
            // Handle the case where _address is undefined or not available
            console.error(
              "Error: Unable to determine address for rows[key1]._cells[13]"
            );
            // Handle this error scenario appropriately
          }
        }
        console.log("Initial Value:", initialValue13);
        console.log("New Value:", newValue13);
  
        if (
          rows[key1]._cells[0] != undefined &&
          rows[key1]._cells[0]._value.model.value == "$$(vc+vp)"
        ) {
          if (
            rows[key1]._cells[11] &&
            rows[key1]._cells[11].value !== undefined
          ) {
            let cell11Value = rows[key1]._cells[11].value;
            let cell11 = rows[key1]._cells[11];
            let cell2 = rows[key1]._cells[2];
            let add11 = cell11._address;
            let add2 = cell2._address;
            Vu_max = rows[key1]._cells[20].value;
            finalResult = pi * (newValue13 + Vp);
            half_finalResult = finalResult / 2;
            console.log(finalResult);
            data = { ...data, [add11]: parseFloat(finalResult.toFixed(2)) };
            data = { ...data, [add2]: parseFloat(half_finalResult.toFixed(2)) };
          } else {
            console.error(
              "Error: Unable to retrieve value for rows[key1]._cells[11]"
            );
          }
        }
        if (
          rows[key1]._cells[0] != undefined &&
          rows[key1]._cells[0]._value.model.value == "$$check"
        ) {
          let cell2 = rows[key1]._cells[2];
          let cell11 = rows[key1]._cells[11];
          let cell12 = rows[key1]._cells[12];
          let add2 = cell2._address;
          let add11 = cell11._address;
          let add12 = cell12._address;
          let cell2Value;
          if (Math.abs(half_finalResult) > Vu_max) {
            cell2Value = "Vu < 0.5Φ(Vc+Vp)";
            data = { ...data, [add2]: "Vu < 0.5Φ(Vc+Vp)" };
            data = { ...data, [add11]: "∴" };
            data = { ...data, [add12]: "No Shear reinforcing" };
          } else {
            cell2Value = "Vu ≥ 0.5ΦVc";
            data = { ...data, [add2]: "Vu ≥ 0.5ΦVc" };
          }
          let key2 = parseInt(key1) + 1;
  
          // Check if rows[key2]._cells[0] value is '$$A,req'
          if (cell2Value == "Vu ≥ 0.5ΦVc") {
            if (
              rows[key2]._cells[0] != undefined &&
              rows[key2]._cells[0]._value.model.value == "$$Ar"
            ) {
              let cell13 = rows[key2]._cells[13];
              let add13 = cell13._address;
              let Av_extra;
              let Avr =
                ((Vu_max - finalResult) * s_max) /
                (pi * fy * dv * (cot(theta_new) + cot(a)) * Math.sin(a));
              console.log(Avr);
              data = { ...data, [add13]: Avr };
              for (let i = key2; i <= worksheet.rowCount; i++) {
                // console.log("Hello");
                let nextRow = worksheet.getRow(i);
                if (
                  rows[nextRow]._cells[0] != undefined &&
                  rows[key1]._cells[0]._value.model.value == "$$Av,req"
                ) {
                  let cell12 = rows[nextRow]._cells[12];
                  let add12 = cell12._address;
                  if (Avm > Avr) {
                    Av_extra = Avm;
                    data = { ...data, [add12]: Av };
                  } else {
                    Av_extra = Avr;
                    data = { ...data, [add12]: Avr };
                  }
                }
                if (
                  rows[nextRow]._cells[0] != undefined &&
                  rows[key1]._cells[0]._value.model.value == "$$A,v"
                ) {
                  let cell11 = rows[nextRow]._cells[11];
                  let cell29 = rows[nextRow]._cells[29];
                  let add11 = cell11._address;
                  let add29 = cell29._address;
                  if (Av >= Av_extra) {
                    data = { ...data, [add11]: "≥" };
                    data = { ...data, [add29]: "OK" };
                  } else {
                    data = { ...data, [add11]: "<" };
                    data = { ...data, [add29]: "NG" };
                  }
                }
                if (nextRow.getCell(1).value === "$$A,v") {
                  // Found $$A,v, break the loop
                  break;
                }
                // Perform your desired operations within the loop here
              }
            } else {
              let key3 = parseInt(key1) + 2;
              // Insert five rows after key3
              // worksheet.spliceRows(key3 + 1, 0, [], [], [], [], []);
  
              // Update references after insertion
              key3 += 5;
  
              let cell19 = rows[key3]._cells[19];
              let add19 = cell19._address;
              data = { ...data, [add19]: "Av,req1" };
              let cell20 = rows[key3]._cells[20];
              let add20 = cell20._address;
              data = { ...data, [add20]: "=" };
              let cell21 = rows[key3]._cells[21];
              let add21 = cell21._address;
              data = { ...data, [add21]: "{ Vu - Φ(Vc+Vp) }·s" };
              let cell21_n = rows[key3 + 1]._cells[21];
              let add21_n = cell21_n._address;
              data = { ...data, [add21_n]: "Φ·fy·dv(cotθ+cotα)sinα" };
              let cell27 = rows[key3]._cells[27];
              let add27 = cell27._address;
              data = { ...data, [add27]: "=" };
              let cell28 = rows[key3]._cells[28];
              let add28 = cell28._address;
              let Av_extra;
              let Avr =
                ((Vu_max - finalResult) * s_max) /
                (pi * fy * dv * (cot(theta_new) + cot(a)) * Math.sin(a));
              data = { ...data, [add28]: Avr };
            }
          } else {
            if (
              rows[key2]._cells[0] != undefined &&
              rows[key2]._cells[0]._value.model.value == "$$Ar"
            ) {
              for (let i = key2 + 1; i <= worksheet.rowCount; i++) {
                let nextRow = worksheet.getRow(i);
  
                if (nextRow.getCell(1).value === "$$A,v") {
                  // First, make the cells blank
                  nextRow.eachCell({ includeEmpty: true }, (cell) => {
                    cell.value = "";
                  });
                  break;
                  // Blank the corresponding rows
                }
                nextRow.eachCell({ includeEmpty: true }, (cell) => {
                  cell.value = "";
                });
              }
            }
          }
        }
        if (
          rows[key1]._cells[0] != undefined &&
          rows[key1]._cells[0]._value.model.value == "$$t_new"
        ) {
          let cell13 = rows[key1]._cells[13];
          let add13 = cell13._address;
          data = { ...data, [add13]: theta_new };
          let cell8 = rows[key1]._cells[8];
          let add8 = cell8._address;
          data = { ...data, [add8]: "########" };
        }
        if (
          rows[key1]._cells[0] != undefined &&
          rows[key1]._cells[0]._value.model.value == "$$sum"
        ) {
          let cell7 = rows[key1]._cells[7];
          let add7 = cell7._address;
          let sum = Vp + Vc + Vs;
          Vn = sum;
          data = { ...data, [add7]: parseFloat(Vn.toFixed(2)) };
          let cell13 = rows[key1]._cells[13];
          let cell19 = rows[key1]._cells[19];
          let add13 = cell13._address;
          let add19 = cell19._address;
          let value19 = cell19.value;
          if (Vn <= value19) {
            data = { ...data, [add13]: "≤" };
          } else {
            data = { ...data, [add13]: ">" };
          }
        }
        if (
          rows[key1]._cells[0] != undefined &&
          rows[key1]._cells[0]._value.model.value == "$$vn"
        ) {
          let cell11 = rows[key1]._cells[11];
          let add11 = cell11._address;
          data = { ...data, [add11]: parseFloat(Vn.toFixed(2)) };
        }
        if (
          rows[key1]._cells[0] != undefined &&
          rows[key1]._cells[0]._value.model.value == "$$vr"
        ) {
          let cell8 = rows[key1]._cells[8];
          let cell17_value = rows[key1]._cells[17].value;
          let cell29 = rows[key1]._cells[29];
          let add29 = cell29._address;
          let add8 = cell8._address;
          let vr = pi * Vn;
          let cell16 = rows[key1]._cells[16];
          let add16 = cell16._address;
          if (vr < cell17_value) {
            data = { ...data, [add16]: "<" };
            data = { ...data, [add29]: "NG" };
          } else {
            data = { ...data, [add16]: "≥" };
            data = { ...data, [add29]: "OK" };
          }
          data = { ...data, [add8]: parseFloat(vr.toFixed(2)) };
        }
  
        if (
          rows[key1]._cells[0] != undefined &&
          rows[key1]._cells[0]._value.model.value == "$$strn"
        ) {
          let add1 = rows[key1]._cells[2]._value.model.address;
          let add2 = rows[key1]._cells[8]._value.model.address;
          data = { ...data, [add1]: "Calculation for β and θ" };
          data = { ...data, [add2]: "" };
          let add12 = rows[key1]._cells[12]._value.model.address;
          data = { ...data,[add12] : '(See CA - 5.7.3.4)'}
        }
        // if (rows[key1]._cells[0] != undefined && rows[key1]._cells[0]._value.model.value == '$$strn1') {
        //     let add1 = rows[key1]._cells[3]._value.model.address;
        //     let add2 = rows[key1]._cells[4]._value.model.address; let add3 = rows[key1]._cells[5]._value.model.address;
        //     data = { ...data, [add1]: 'β' };
        //     data = { ...data, [add2]: '=' };
        //     data = { ...data, [add3]: beta2.toFixed(3) };
  
        //     let add4 = rows[key1]._cells[8]._value.model.address;
        //     let add5 = rows[key1]._cells[9]._value.model.address;
        //     let add6 = rows[key1]._cells[10]._value.model.address;
        //     data = { ...data, [add4]: 'θ' };
        //     data = { ...data, [add5]: '=' };
        //     data = { ...data, [add6]: theta2.toFixed(3) };
  
        //     let add7 = rows[key1]._cells[13]._value.model.address;
        //     let add8 = rows[key1]._cells[14]._value.model.address;
        //     let add9 = rows[key1]._cells[15]._value.model.address;
        //     data = { ...data, [add7]: 'εₓ' };
        //     data = { ...data, [add8]: '=' };
        //     data = { ...data, [add9]: Exn.toFixed(8) };
        // }
        // if (rows[key1]._cells[0] != undefined && rows[key1]._cells[0]._value.model.value == '$$vc') {
        //     let add1 = rows[key1]._cells[9]._value.model.address;
        //     data = { ...data, [add1]: Vc1 };
        // }
        if (
          rows[key1]._cells[0] != undefined &&
          rows[key1]._cells[0]._value.model.value == "$$A_min"
        ) {
          console.log(rows[key1]._cells);
          let cell = rows[key1]._cells[4];
          if (cell && cell._address) {
            let add11 = cell._address;
            data = { ...data, [add11]: "A" };
          } else {
            console.error(
              "Error: Unable to determine address for rows[key1]._cells[4]"
            );
          }
          let cell2 = rows[key1]._cells[6];
          if (cell2 && cell2._address) {
            let add12 = cell2._address;
            data = { ...data, [add12]: "Aₘᵢₙ" };
          } else {
            console.error(
              "Error: Unable to determine address for rows[key1]._cells[4]"
            );
          }
          let cell3 = rows[key1]._cells[5];
          let comparisonSymbol = Av >= Avm ? "≥" : "<";
          if (cell3 && cell3._address) {
            let add13 = cell3._address;
            data = { ...data, [add13]: comparisonSymbol };
          }
        }
        if (
          rows[key1]._cells[0] != undefined &&
          rows[key1]._cells[0]._value.model.value == "$$e_min"
        ) {
          console.log(rows[key1]._cells);
          let cell = rows[key1]._cells[4];
          if (cell && cell._address) {
            let add11 = cell._address;
            data = { ...data, [add11]: "εx" };
          } else {
            // Handle the case where _address is undefined or not available
            console.error(
              "Error: Unable to determine address for rows[key1]._cells[4]"
            );
            // Handle this error scenario appropriately
          }
          let cell2 = rows[key1]._cells[5];
  
          // Check if cell and its properties are defined
          if (cell2 && cell2._address) {
            let add12 = cell2._address;
            data = { ...data, [add12]: "=" };
          } else {
            // Handle the case where _address is undefined or not available
            console.error(
              "Error: Unable to determine address for rows[key1]._cells[4]"
            );
            // Handle this error scenario appropriately
          }
          let cell3 = rows[key1]._cells[6];
          if (cell3 && cell3._address) {
            let add13 = cell3._address;
            let cell3Value = Av >= Avm ? Exn : Etn;
            data = { ...data, [add13]: cell3Value };
          }
        }
        if (
          rows[key1]._cells[0] != undefined &&
          rows[key1]._cells[0]._value.model.value == "$$sx_min"
        ) {
          rows[key1]._cells = rows[key1]._cells.map((cell) =>
            cell === "" ? undefined : cell
          );
          console.log(rows[key1]._cells);
          let cell = rows[key1]._cells[4];
          if (cell && cell._address) {
            let add11 = cell._address;
            data = { ...data, [add11]: "sx" };
          } else {
            console.error(
              "Error: Unable to determine address for rows[key1]._cells[4]"
            );
          }
  
          let cell2 = rows[key1]._cells[5];
          if (cell2 && cell2._address) {
            let add12 = cell2._address;
            data = { ...data, [add12]: "=" };
          } else {
            console.error(
              "Error: Unable to determine address for rows[key1]._cells[5]"
            );
          }
  
          let startCol = 7;
          let endCol = 17;
          let rowNumber = rows[key1]._cells[startCol].row;
          try {
            let mergeRange = worksheet.getCell(
              `${worksheet.getColumn(startCol).letter}${rowNumber}:${
                worksheet.getColumn(endCol).letter
              }${rowNumber}`
            );
            if (!mergeRange.isMerged) {
              worksheet.mergeCells(rowNumber, startCol, rowNumber, endCol);
            }
          } catch (error) {
            console.error("Error merging cells: ", error);
          }
  
          let cell3 = worksheet.getCell(rowNumber, startCol);
          if (cell3 && cell3._address) {
            let add13 = cell3._address;
  
            if (Av < Avm) {
              data = {
                ...data,
                [add13]: `Min| dv, maximum distance between the longitudinal r/f |`,
              };
              let cell4 = rows[key1]._cells[18];
              if (cell4 && cell4._address) {
                let add14 = cell4._address;
                data = { ...data, [add14]: "=" };
              } else {
                console.error(
                  "Error: Unable to determine address for rows[key1]._cells[24]"
                );
              }
  
              let cell5 = rows[key1]._cells[19];
              if (cell5 && cell5._address) {
                let add15 = cell5._address;
                let add15value = dv < sg ? dv : sg;
                data = { ...data, [add15]: add15value };
              } else {
                console.error(
                  "Error: Unable to determine address for rows[key1]._cells[25]"
                );
              }
            } else {
              data = { ...data, [add13]: `Not Required` };
            }
          } else {
            console.error(
              "Error: Unable to determine address for rows[key1]._cells[24]"
            );
          }
        }
  
        if (
          rows[key1]._cells[0] != undefined &&
          rows[key1]._cells[0]._value.model.value == "$$sxe_min"
        ) {
          let cell3 = rows[key1]._cells[4];
          if (cell3 && cell3._address) {
            let add13 = cell3._address;
            // let sxe = ((add15value*1.38)/(ag +0.63))
            data = { ...data, [add13]: "sxe" };
          }
          let cell4 = rows[key1]._cells[5];
          if (cell4 && cell4._address) {
            let add14 = cell4._address;
            // let sxe = ((add15value*1.38)/(ag +0.63))
            data = { ...data, [add14]: "=" };
          }
          let mergeStartCol = 7;
          let mergeEndCol = 16;
          let mergeRowNumber = rows[key1]._cells[mergeStartCol].row;
  
          // Check if the range is already merged
          try {
            let mergeRange = worksheet.getCell(
              `${worksheet.getColumn(mergeStartCol).letter}${mergeRowNumber}:${
                worksheet.getColumn(mergeEndCol).letter
              }${mergeRowNumber}`
            );
            if (!mergeRange.isMerged) {
              worksheet.mergeCells(
                mergeRowNumber,
                mergeStartCol,
                mergeRowNumber,
                mergeEndCol
              );
            }
          } catch (error) {
            console.error("Error merging cells: ", error);
          }
  
          // After merging, the cell5 should refer to the merged cell
          let cell5 = worksheet.getCell(mergeRowNumber, mergeStartCol);
  
          if (cell5 && cell5._address) {
            let add15 = cell5._address;
  
            if (Av < Avm) {
              data = { ...data, [add15]: sxe };
            } else {
              data = { ...data, [add15]: "Not required" };
            }
          } else {
            console.error(
              "Error: Unable to determine address for rows[key1]._cells[6]"
            );
          }
        }
  
        if (
          rows[key1]._cells[0] != undefined &&
          rows[key1]._cells[0]._value.model.value == "$$b_min"
        ) {
          let cell5 = rows[key1]._cells[4];
          let add5 = cell5._address;
          data = { ...data, [add5]: "β" };
          let cell6 = rows[key1]._cells[5];
          let add6 = cell6._address;
          data = { ...data, [add6]: "=" };
          let cell7 = rows[key1]._cells[6];
          let add7 = cell7._address;
          if (Av < Avm) {
            let b_value = ThetaBeta2(sxe, Etn * 1000);
            let beta = parseFloat(b_value[0].toFixed(2));
            console.log(beta);
            data = { ...data, [add7]: beta };
            beta_new_min = beta;
          } else {
            data = { ...data, [add7]: parseFloat(beta2.toFixed(2)) };
            beta_new_min = parseFloat(beta2.toFixed(2));
          }
          let cell15 = rows[key1]._cells[15];
          let add15 = cell15._address;
          data = { ...data,[add15] : '(See CA-5.7.3.4)'}
        }
  
        if (
          rows[key1]._cells[0] != undefined &&
          rows[key1]._cells[0]._value.model.value == "$$theta_min"
        ) {
          let cell5 = rows[key1]._cells[4];
          let add5 = cell5._address;
          data = { ...data, [add5]: "θ" };
          let cell6 = rows[key1]._cells[5];
          let add6 = cell6._address;
          data = { ...data, [add6]: "=" };
          let cell7 = rows[key1]._cells[6];
          let add7 = cell7._address;
          if (Av < Avm) {
            let theta_value = ThetaBeta2(sxe, Etn * 1000);
            let theta = parseFloat(theta_value[1].toFixed(2));
            console.log(theta);
            data = { ...data, [add7]: theta };
            theta_new_min = theta;
          } else {
            data = { ...data, [add7]: parseFloat(theta2.toFixed(2)) };
            theta_new_min = parseFloat(theta2.toFixed(2));
          }
          let cell15 = rows[key1]._cells[15];
          let add15 = cell15._address;
          data = { ...data,[add15] : '(See CA-5.7.3.4)'}
        }
  
        if (
          rows[key1]._cells[0] != undefined &&
          rows[key1]._cells[0]._value.model.value == "$$beta_min"
        ) {
          let cell9 = rows[key1]._cells[9];
          if (cell9 && cell9._address) {
            beta = cell9._value.model.value;
          } else {
            // Handle the case where _address is undefined or not available
            console.error(
              "Error: Unable to determine address for rows[key1]._cells[9]"
            );
            // Handle this error scenario appropriately
          }
        }
        if (
          rows[key1]._cells[0] != undefined &&
          rows[key1]._cells[0]._value.model.value == "$$vc_min"
        ) {
          let cell13 = rows[key1]._cells[13];
          let cell9 = rows[key1]._cells[9];
          let cell5 = rows[key1]._cells[5];
          let add5 = cell5._address;
          let add9 = cell9._address;
          data = { ...data, [add5]: "0.0316" };
          console.log(rows[key1]._cells[9].value);
          console.log(rows[key1]);
          if (cell13 && cell13._address) {
            if (type == "Composite") {
              let add13 = cell13._address;
              // Store the initial value globally
              initialValue13 = cell13._value.model.value;
              // Retrieve the value from cell13, divide it by beta, and multiply by beta1
              let value13 = initialValue13;
              let result = (value13 / beta) * beta1;
              // Store the result back in cell13
              cell13._value.model.value = result;
              // Store the new value globally
              newValue13 = result; //new Vc value
              data = { ...data, [add9]: "β √f'c bvdv" };
              data = { ...data, [add13]: parseFloat(result.toFixed(2)) };
              Vc_min = newValue13;
            } else {
              let add13 = cell13._address;
              // Store the initial value globally
              initialValue13 = cell13._value.model.value;
              // Retrieve the value from cell13, divide it by beta, and multiply by beta1
              let value13 = initialValue13;
              let result = (value13 / K) * beta1;
              // Store the result back in cell13
              cell13._value.model.value = result;
              // Store the new value globally
              newValue13 = result; //new Vc value
              data = { ...data, [add9]: "β √f'c bvdv" };
              data = { ...data, [add13]: parseFloat(result.toFixed(2)) };
              Vc_min = newValue13;
            }
          } else {
            // Handle the case where _address is undefined or not available
            console.error(
              "Error: Unable to determine address for rows[key1]._cells[13]"
            );
            // Handle this error scenario appropriately
          }
        }
        // console.log("Initial Value:", initialValue13);
        // console.log("New Value:", newValue13);
  
        if (
          rows[key1]._cells[0] != undefined &&
          rows[key1]._cells[0]._value.model.value == "$$(vc+vp)_min"
        ) {
          if (
            rows[key1]._cells[11] &&
            rows[key1]._cells[11].value !== undefined
          ) {
            let cell11Value = rows[key1]._cells[11].value;
            let cell11 = rows[key1]._cells[11];
            let cell2 = rows[key1]._cells[2];
            let add11 = cell11._address;
            let add2 = cell2._address;
            Vu_min = rows[key1]._cells[20].value;
            finalResult = pi * (Vc_min + Vp);
            half_finalResult = finalResult / 2;
            console.log(finalResult);
            data = { ...data, [add11]: parseFloat(finalResult.toFixed(2)) };
            data = { ...data, [add2]: parseFloat(half_finalResult.toFixed(2)) };
          } else {
            console.error(
              "Error: Unable to retrieve value for rows[key1]._cells[11]"
            );
          }
        }
        if (
          rows[key1]._cells[0] != undefined &&
          rows[key1]._cells[0]._value.model.value == "$$check_min"
        ) {
          let cell2 = rows[key1]._cells[2];
          let cell11 = rows[key1]._cells[11];
          let cell12 = rows[key1]._cells[12];
          let add2 = cell2._address;
          let add11 = cell11._address;
          let add12 = cell12._address;
          let cell2Value;
          if (Math.abs(half_finalResult) > Vu_min) {
            data = { ...data, [add2]: "Vu < 0.5Φ(Vc+Vp)" };
            data = { ...data, [add11]: "∴" };
            data = { ...data, [add12]: "No Shear reinforcing" };
          } else {
            data = { ...data, [add2]: "Vu ≥ 0.5ΦVc" };
          }
          let key2 = parseInt(key1) + 1;
  
          // Check if rows[key2]._cells[0] value is '$$A,req'
          if (cell2Value == "Vu ≥ 0.5ΦVc") {
            if (
              rows[key2]._cells[0] != undefined &&
              rows[key2]._cells[0]._value.model.value == "$$Ar_min"
            ) {
              let cell13 = rows[key2]._cells[13];
              let add13 = cell13._address;
              let Av_extra;
              let Avr =
                ((Vu_min - finalResult) * s_min) /
                (pi * fy * dv_min * (cot(theta_new_min) + cot(a)) * Math.sin(a));
              console.log(Avr);
              data = { ...data, [add13]: Avr };
              for (let i = key2; i <= worksheet.rowCount; i++) {
                // console.log("Hello");
                let nextRow = worksheet.getRow(i);
                if (
                  rows[nextRow]._cells[0] != undefined &&
                  rows[key1]._cells[0]._value.model.value == "$$Av,req_min"
                ) {
                  let cell12 = rows[nextRow]._cells[12];
                  let add12 = cell12._address;
                  if (Avm > Avr) {
                    Av_extra = Avm;
                    data = { ...data, [add12]: Av };
                  } else {
                    Av_extra = Avr;
                    data = { ...data, [add12]: Avr };
                  }
                }
                if (
                  rows[nextRow]._cells[0] != undefined &&
                  rows[key1]._cells[0]._value.model.value == "$$A,v_min"
                ) {
                  let cell11 = rows[nextRow]._cells[11];
                  let cell29 = rows[nextRow]._cells[29];
                  let add11 = cell11._address;
                  let add29 = cell29._address;
                  if (Av >= Av_extra) {
                    data = { ...data, [add11]: "≥" };
                    data = { ...data, [add29]: "OK" };
                  } else {
                    data = { ...data, [add11]: "<" };
                    data = { ...data, [add29]: "NG" };
                  }
                }
                if (nextRow.getCell(1).value === "$$A,v") {
                  // Found $$A,v, break the loop
                  break;
                }
                // Perform your desired operations within the loop here
              }
            } else {
              let key3 = parseInt(key1) + 2;
              if (
                rows[key3]._cells[0] != undefined &&
                rows[key3]._cells[0]._value.model.value == "$$vs_min"
              ) {
                let cell19 = rows[key3]._cells[19];
                let add19 = cell19._address;
                data = { ...data, [add19]: "Av,req1" };
                let cell20 = rows[key3]._cells[20];
                let add20 = cell20._address;
                data = { ...data, [add20]: "=" };
                let cell21 = rows[key3]._cells[21];
                let add21 = cell21._address;
                data = { ...data, [add21]: "{ Vu - Φ(Vc+Vp) }·s" };
                let cell21_n = rows[key3 + 1]._cells[21];
                let add21_n = cell21_n._address;
                data = { ...data, [add21_n]: "Φ·fy·dv(cotθ+cotα)sinα" };
                let cell27 = rows[key3]._cells[27];
                let add27 = cell27._address;
                data = { ...data, [add27]: "=" };
                let cell28 = rows[key3]._cells[28];
                let add28 = cell28._address;
                let Av_extra;
                let Avr =
                  ((Vu_max - finalResult) * s_max) /
                  (pi * fy * dv * (cot(theta_new) + cot(a)) * Math.sin(a));
                data = { ...data, [add28]: Avr };
              }
            }
          } else {
            if (
              rows[key2]._cells[0] != undefined &&
              rows[key2]._cells[0]._value.model.value == "$$Ar_min"
            ) {
              for (let i = key2 + 1; i <= worksheet.rowCount; i++) {
                let nextRow = worksheet.getRow(i);
  
                if (nextRow.getCell(1).value === "$$A,v") {
                  // First, make the cells blank
                  nextRow.eachCell({ includeEmpty: true }, (cell) => {
                    cell.value = "";
                  });
                  break;
                  // Blank the corresponding rows
                }
                nextRow.eachCell({ includeEmpty: true }, (cell) => {
                  cell.value = "";
                });
              }
            }
          }
        }
        if (
          rows[key1]._cells[0] != undefined &&
          rows[key1]._cells[0]._value.model.value == "$$t_new_min"
        ) {
          let cell13 = rows[key1]._cells[13];
          let add13 = cell13._address;
          data = { ...data, [add13]: theta_new_min };
          let cell8 = rows[key1]._cells[8];
          let add8 = cell8._address;
          data = { ...data, [add8]: "########" };
        }
        if (
          rows[key1]._cells[0] != undefined &&
          rows[key1]._cells[0]._value.model.value == "$$sum_min"
        ) {
          let cell7 = rows[key1]._cells[7];
          let add7 = cell7._address;
          let sum = Vp_min + Vc_min + Vs_min;
          Vn_min = sum;
          data = { ...data, [add7]: parseFloat(Vn_min.toFixed(2)) };
          let cell13 = rows[key1]._cells[13];
          let cell19 = rows[key1]._cells[19];
          let add13 = cell13._address;
          let add19 = cell19._address;
          let value19 = cell19.value;
          if (Vn_min <= value19) {
            data = { ...data, [add13]: "≤" };
          } else {
            data = { ...data, [add13]: ">" };
          }
        }
        if (
          rows[key1]._cells[0] != undefined &&
          rows[key1]._cells[0]._value.model.value == "$$vn"
        ) {
          let cell11 = rows[key1]._cells[11];
          let add11 = cell11._address;
          data = { ...data, [add11]: parseFloat(Vn.toFixed(2)) };
        }
        if (
          rows[key1]._cells[0] != undefined &&
          rows[key1]._cells[0]._value.model.value == "$$vn_min"
        ) {
          let cell11 = rows[key1]._cells[11];
          let add11 = cell11._address;
          data = { ...data, [add11]: parseFloat(Vn_min.toFixed(2)) };
        }
        if (
          rows[key1]._cells[0] != undefined &&
          rows[key1]._cells[0]._value.model.value == "$$vr"
        ) {
          let cell8 = rows[key1]._cells[8];
          let cell17_value = rows[key1]._cells[17].value;
          let cell29 = rows[key1]._cells[29];
          let add29 = cell29._address;
          let add8 = cell8._address;
          let vr = pi * Vn;
          let cell16 = rows[key1]._cells[16];
          let add16 = cell16._address;
          if (vr < cell17_value) {
            data = { ...data, [add16]: "<" };
            data = { ...data, [add29]: "NG" };
          } else {
            data = { ...data, [add16]: "≥" };
            data = { ...data, [add29]: "OK" };
          }
          data = { ...data, [add8]: parseFloat(vr.toFixed(2)) };
        }
        if (
          rows[key1]._cells[0] != undefined &&
          rows[key1]._cells[0]._value.model.value == "$$vr_min"
        ) {
          let cell8 = rows[key1]._cells[8];
          let cell17_value = rows[key1]._cells[17].value;
          let cell29 = rows[key1]._cells[29];
          let add29 = cell29._address;
          let add8 = cell8._address;
          let vr_min = pi_min * Vn_min;
          let cell16 = rows[key1]._cells[16];
          let add16 = cell16._address;
          if (vr_min < cell17_value) {
            data = { ...data, [add16]: "<" };
            data = { ...data, [add29]: "NG" };
          } else {
            data = { ...data, [add16]: "≥" };
            data = { ...data, [add29]: "OK" };
          }
          data = { ...data, [add8]: parseFloat(vr_min.toFixed(2)) };
        }
        function cot(angle) {
          return 1 / Math.tan(angle);
        }
  
        if (
          rows[key1]._cells[0] != undefined &&
          rows[key1]._cells[0]._value.model.value == "$$vs"
        ) {
          let cell13 = rows[key1]._cells[13];
          let add13 = cell13._address;
          let cal =
            (Av * fy * dv * (cot(theta_new) + cot($$alpha)) * Math.sin($$alpha)) /
            s_max;
          data = { ...data, [add13]: cal };
          Vs = cal;
        }
        if (
          rows[key1]._cells[0] != undefined &&
          rows[key1]._cells[0]._value.model.value == "$$vs_min"
        ) {
          let cell13 = rows[key1]._cells[13];
          let add13 = cell13._address;
          let cal =
            (Av *
              fy *
              dv_min *
              (cot(theta_new_min) + cot($$alpha_min)) *
              Math.sin($$alpha_min)) /
            s_min;
          data = { ...data, [add13]: cal };
          Vs_min = cal;
        }
      }
  
      for (let key in data) {
        let match = key.match(/^([A-Za-z]+)(\d+)$/);
        if (match) {
          const row = match[1];
          const col = match[2];
          let value = 0;
          let factor = 1;
          for (let i = row.length - 1; i >= 0; i--) {
            value += (row.charCodeAt(i) - 64) * factor;
            factor *= 26;
          }
          worksheet._rows[col - 1]._cells[value - 1]._value.model.value =
            data[key];
          worksheet._rows[col - 1]._cells[value - 1]._value.model.type = 3;
        }
      }
      for (let key1 in rows) {
        if (
          rows[key1]._cells[0] != undefined &&
          rows[key1]._cells[0]._value.model.value == "$$b_str"
        ) {
          // Store the starting index for deletion
          let startIdx = parseInt(key1);
          let rownumber = 0;
          let endIdx;
  
          // Loop through rows starting from the next row after '$$b_str'
          for (let i = startIdx; i < rows.length; i++) {
            // Check if the current row has the value '$$b_end'
            if (
              rows[i] &&
              rows[i]._cells[0] &&
              rows[i]._cells[0]._value.model.value == "$$b_end"
            ) {
              // Store the ending index for deletion
              endIdx = i;
              // Calculate the number of rows to delete
              rownumber = endIdx - startIdx + 1;
              // Delete rows between '$$b_str' and '$$b_end' (inclusive)
              worksheet._rows.splice(startIdx, rownumber);
              break;
            }
          }
          // Remove the old references
          worksheet._rows.length -= rownumber;
  
          break;
        }
        // if (
        //   rows[key1]._cells[0] != undefined &&
        //   rows[key1]._cells[0]._value.model.value == "$$b_str_min"
        // ) {
        //   // Store the starting index for deletion
        //   let startIdx = parseInt(key1);
        //   let rownumber = 0;
        //   let endIdx;
  
        //   // Loop through rows starting from the next row after '$$b_str'
        //   for (let i = startIdx; i < rows.length; i++) {
        //     // Check if the current row has the value '$$b_end'
        //     if (
        //       rows[i] &&
        //       rows[i]._cells[0] &&
        //       rows[i]._cells[0]._value.model.value == "$$b_end_min"
        //     ) {
        //       // Store the ending index for deletion
        //       endIdx = i;
        //       // Calculate the number of rows to delete
        //       rownumber = endIdx - startIdx + 1;
        //       // Delete rows between '$$b_str' and '$$b_end' (inclusive)
        //       worksheet._rows.splice(startIdx, rownumber);
        //       break;
        //     }
        //   }
        // }
        // const pattern = /^\$\$.*/; // Regular expression to match '$$' followed by any characters   
        
      }
      for (let key1 in rows) {
        if (
          rows[key1]._cells[0] != undefined &&
          rows[key1]._cells[0]._value.model.value == "$$b_str_min"
        ) {
          // Store the starting index for deletion
          let startIdx = parseInt(key1);
          let rownumber = 0;
          let endIdx;
  
          // Loop through rows starting from the next row after '$$b_str'
          for (let i = startIdx; i < rows.length; i++) {
            // Check if the current row has the value '$$b_end'
            if (
              rows[i] &&
              rows[i]._cells[0] &&
              rows[i]._cells[0]._value.model.value == "$$b_end_min"
            ) {
              // Store the ending index for deletion
              endIdx = i;
              // Calculate the number of rows to delete
              rownumber = endIdx - startIdx + 1;
              // Delete rows between '$$b_str' and '$$b_end' (inclusive)
              worksheet._rows.splice(startIdx, rownumber);
              break;
            }
          }
          // Remove the old references
          worksheet._rows.length -= rownumber;
  
          break;
        }
        // if (
        //   rows[key1]._cells[0] != undefined &&
        //   rows[key1]._cells[0]._value.model.value == "$$b_str_min"
        // ) {
        //   // Store the starting index for deletion
        //   let startIdx = parseInt(key1);
        //   let rownumber = 0;
        //   let endIdx;
  
        //   // Loop through rows starting from the next row after '$$b_str'
        //   for (let i = startIdx; i < rows.length; i++) {
        //     // Check if the current row has the value '$$b_end'
        //     if (
        //       rows[i] &&
        //       rows[i]._cells[0] &&
        //       rows[i]._cells[0]._value.model.value == "$$b_end_min"
        //     ) {
        //       // Store the ending index for deletion
        //       endIdx = i;
        //       // Calculate the number of rows to delete
        //       rownumber = endIdx - startIdx + 1;
        //       // Delete rows between '$$b_str' and '$$b_end' (inclusive)
        //       worksheet._rows.splice(startIdx, rownumber);
        //       break;
        //     }
        //   }
        // }
        // const pattern = /^\$\$.*/; // Regular expression to match '$$' followed by any characters   
        
      }
      for (let i = 0; i < rows.length; i++) {
            if (rows[i] && rows[i]._cells && rows[i]._cells[0]) {
            const cellValue = rows[i]._cells[0]?._value?.model?.value;
        
            // Log the cell value for debugging
            console.log(`Row ${i} cell value:`, safeStringify(cellValue));
        
            // Ensure cellValue is a string
            if (cellValue && typeof cellValue === 'string') {
                const charArray = Array.from(cellValue);
        
                if (charArray[0] === '$' && charArray[1] === '$') {
                    rows[i]._cells[0].value = ''; // Make the cell blank
                }
            }
        
            const row = worksheet.getRow(i + 1); // Adjust index based on your worksheet API
            row.getCell(1).value = rows[i]._cells[0]?.value || ''; // Update the cell with the new value or an empty string
            row.commit(); // Commit the changes (if needed by the library you are using)
        } else {
            // Log or handle the case where rows[i] or _cells[0] is undefined
            console.log(`Skipping row ${i} due to undefined or missing _cells property.`);
        }
    }
      workbookData.worksheets[wkey] = worksheet;
      setWorkbookData(workbookData);
      setSheetName(worksheet.name);
    }
  
    function updatedata2(wkey, worksheet2,beamStresses) {
      if (!workbookData) return;
      if (!worksheet2) {
        throw new Error("No worksheets found in the uploaded file");
      }
        // Get the number of rows in the worksheet
        const rowCount = worksheet2.rowCount;

        // Access the last row
        const lastRowNumber = rowCount; // Row numbers are 1-based
        const lastRow = worksheet2.getRow(lastRowNumber);

        // Log the last row for debugging
        console.log(`Last row (${lastRowNumber}):`, lastRow);
        const nextRowNumber = lastRowNumber + 1;

        // Access the next row
        const nextRow = worksheet2.getRow(nextRowNumber);
    
        // Populate the first cell with selectedName
        nextRow.getCell(1).value = beamStresses.BeamStress.DATA[0][1];
        nextRow.getCell(2).value = beamStresses.BeamStress.DATA[0][5];
        nextRow.getCell(3).value = 'Girder';
        nextRow.getCell(4).value = 'Tension';
        nextRow.getCell(5).value = selectedName;
        nextRow.getCell(8).value = changeSignAndFormat(beamStresses.BeamStress.DATA[0][12]);
        nextRow.getCell(9).value = changeSignAndFormat(beamStresses.BeamStress.DATA[0][15]);
        nextRow.getCell(10).value = changeSignAndFormat(beamStresses.BeamStress.DATA[0][13]);
        nextRow.getCell(11).value = changeSignAndFormat(beamStresses.BeamStress.DATA[0][14]); 
        nextRow.getCell(6).value = calculateAverage(nextRow.getCell(8).value, nextRow.getCell(10).value);
        nextRow.getCell(7).value = calculateAverage(nextRow.getCell(9).value, nextRow.getCell(11).value);
        nextRow.getCell(12).value = findMinValue([ nextRow.getCell(8),nextRow.getCell(9),nextRow.getCell(10),nextRow.getCell(11)]);
        nextRow.getCell(13).value = '0';
        if (nextRow.getCell(12).value > 0) {
            nextRow.getCell(14).value = 'OK';
        }
        else {
            nextRow.getCell(14).value = 'NG';
        }
        // Populate the second cell with Section part from beamStresses.data
        // Assuming beamStresses.data is an array and we want the first element
    
        // Log the next row for debugging
        console.log(`Next row (${nextRowNumber}):`, nextRow);

        // Perform operations on the last row
        // For example, you can get cell values, update them, etc.
        // const lastCellValue = lastRow.getCell(1).value;
        // lastRow.getCell(1).value = 'New Value';

        // Save changes to the last row (if necessary)
        const nextRowNumber2 = lastRowNumber + 2;

        // Access the next row
        const nextRow2 = worksheet2.getRow(nextRowNumber2);
    
        // Populate the first cell with selectedName
        nextRow2.getCell(1).value = beamStresses.BeamStress.DATA[1][1];
        nextRow2.getCell(2).value = beamStresses.BeamStress.DATA[1][5];
        nextRow2.getCell(3).value = 'Slab';
        nextRow2.getCell(4).value = 'Tension';
        nextRow2.getCell(5).value = selectedName;
        nextRow2.getCell(8).value = changeSignAndFormat(beamStresses.BeamStress.DATA[1][12]);
        nextRow2.getCell(9).value = changeSignAndFormat(beamStresses.BeamStress.DATA[1][15]);
        nextRow2.getCell(10).value = changeSignAndFormat(beamStresses.BeamStress.DATA[1][13]);
        nextRow2.getCell(11).value = changeSignAndFormat(beamStresses.BeamStress.DATA[1][14]); 
        nextRow2.getCell(6).value = calculateAverage(nextRow2.getCell(8).value, nextRow2.getCell(10).value);
        nextRow2.getCell(7).value = calculateAverage(nextRow2.getCell(9).value, nextRow2.getCell(11).value);
        nextRow2.getCell(12).value = findMinValue([ nextRow2.getCell(8),nextRow2.getCell(9),nextRow2.getCell(10),nextRow2.getCell(11)]);
        nextRow2.getCell(13).value = '0';
        if (nextRow2.getCell(12).value > 0) {
            nextRow2.getCell(14).value = 'OK';
        }
        else {
            nextRow2.getCell(14).value = 'NG';
        }
        // Populate the second cell with Section part from beamStresses.data
        // Assuming beamStresses.data is an array and we want the first element
    
        // Log the next row for debugging
        console.log(`Next row (${nextRowNumber2}):`, nextRow2);
        lastRow.commit();
    }
    function safeStringify(obj) {
        const cache = new Set();
        return JSON.stringify(obj, (key, value) => {
            if (typeof value === 'object' && value !== null) {
                if (cache.has(value)) {
                    return; // Duplicate reference found, discard key
                }
                cache.add(value); // Store value in our collection
            }
            return value;
        });
    }
    function changeSignAndFormat(value) {
        // Convert the value to a number, change its sign, and format it to two decimal points
        return (-parseFloat(value)).toFixed(2);
    }

    function calculateAverage(value1, value2) {
        return ((parseFloat(value1) + parseFloat(value2)) / 2).toFixed(2);
    }
    function findMinValue(cells) {
        return Math.min(...cells.map(cell => parseFloat(cell.value))).toFixed(2);
    }
    
    function updatedata3(wkey, worksheet3) {
      if (!workbookData) return;
      if (!worksheet3) {
        throw new Error("No worksheets found in the uploaded file");
      }
      console.log(mu_pos);
      let rows = worksheet3._rows;
  
    
      for (let i = 0; i < rows.length; i++) {
        let firstCell = rows[i].getCell(1); // Assuming 1-based indexing for cells
        if (firstCell.value === 'Mu') {
          console.log(`Row ${i + 1} has 'MU' in the first cell.`);
          
          // Change the value of cell 2 to mu_pos
          let secondCell = rows[i].getCell(2);
          secondCell.value = mu_pos; // Update this to the actual value you want to set
          
          // Commit the row changes if necessary
          rows[i].commit();
        }
      }
    }
  
    async function fetchLc() {
      const endpointsDataKeys = [
        { endpoint: "/db/lcom-gen", dataKey: "LCOM-GEN" },
        { endpoint: "/db/lcom-conc", dataKey: "LCOM-CONC" },
        { endpoint: "/db/lcom-src", dataKey: "LCOM-SRC" },
        { endpoint: "/db/lcom-steel", dataKey: "LCOM-STEEL" },
        { endpoint: "/db/lcom-stlcomp", dataKey: "LCOM-STLCOMP" },
      ];
      let allData = [];
      let check = false;
  
      // try {
      for (const { endpoint, dataKey } of endpointsDataKeys) {
        const response = await midasAPI("GET", endpoint);
        console.log(response);
        if (response && !response.error) {
          let responseData = response[dataKey];
          if (responseData === undefined) {
            console.warn(`Data from ${endpoint} is undefined, skipping.`);
            continue;
          }
          if (!Array.isArray(responseData)) {
            responseData = Object.values(responseData);
          }
          const keys = Object.keys(response[dataKey]);
          const lastindex = parseInt(keys[keys.length - 1]);
          console.log(lastindex);
          responseData.forEach((item) => {
            allData.push({ name: item.NAME, endpoint, lastindex: lastindex });
          });
          if (allData.length > 0) {
            const lastElement = allData[0];
            const lastNumber = Object.keys(lastElement).length - 1;
            for (let index = 0; index < responseData.length; index++) {
              const item = responseData[index];
              item.someProperty = lastNumber + index + 1;
              allData.push(item);
              console.log(allData);
            }
          } else {
            allData = allData.concat(responseData);
            console.log(allData);
          }
          check = true;
          console.log(`Data from ${endpoint}:`, responseData);
        }
      }
  
      if (check) {
        setLc(allData);
        // return null;
      }
      showLc(allData);
      // } catch (error) {
      //   enqueueSnackbar("Unable to Fetch Data Check Connection", {
      //     variant: "error",
      //     anchorOrigin: {
      //       vertical: "top",
      //       horizontal: "center",
      //     },
      //   });
      return null;
    }
    function showLc(lc) {
      console.log(lc);
      item.delete("1"); // Make sure 'items' is defined in your context
      let newKey = 1;
      for (let key in lc) {
        if (lc[key].ACTIVE === "SERVICE") {
          item.set(lc[key].NAME, newKey.toString());
          newKey++;
          // console.log(key, lc[key].NAME)
        }
      }
  
      setItem(item);
      // console.log(item);
      // Set default selectedName based on key value 1
      for (let [name, key] of item.entries()) {
        if (key === "1") {
          setSelectedName(name);
          lcname = name;
          console.log(lcname);
          break;
        }
      }
      console.log(item);
    }
    console.log(item);
    console.log(lcname);
    console.log(matchedParts);
    const handleFileDownload = async () => {
      // fetchLc();
      const combArray = Object.values(lc);
      let beamStresses;
      if (combArray.length === 0) {
        enqueueSnackbar("Please Define Load Combination", {
          variant: "error",
          anchorOrigin: {
            vertical: "top",
            horizontal: "center",
          },
        });
        return;
      }
      // console.log('load combinations', Lc)
      console.log(SelectWorksheets);
      console.log(SelectWorksheets2);
      let numberPart = parseInt(matchedParts[0].numberPart, 10); // Extract numberPart from matchParts
      let letterPart = matchedParts[0].letterPart;
      console.log(selectedName)
      const concatenatedValue = `${selectedName}(CBC)`;
      const concatenatedValue_max = `${selectedName}(CBC:max)`;
      let stresses = {
        "Argument": {
            "TABLE_NAME": "BeamStress",
            "TABLE_TYPE": "COMPSECTBEAMSTRESS",
            "EXPORT_PATH": "C:\\MIDAS\\Result\\Output.JSON",
            "STYLES": {
                "FORMAT": "Fixed",
                "PLACE": 12
            },
            "COMPONENTS": [
                "Elem",
                "DOF",
                "Load",
                "Section Part",
                "Part",
                "Axial",
                "Bend(+y)",
                "Bend(-y)",
                "Bend(+z)",
                "Bend(-z)",
                "Cb(min/max)",
                "Cb1(-y+z)",
                "Cb2(+y+z)",
                "Cb3(+y-z)",
                "Cb4(-y-z)",
                "Sax(Warping)1",
                "Sax(Warping)2",
                "Sax(Warping)3",
                "Sax(Warping)4"
            ],
            "NODE_ELEMS": {
                "KEYS": [
                    numberPart
                ]
            },
            "LOAD_CASE_NAMES": [
                selectedName,
                concatenatedValue,
                concatenatedValue_max
            ],
            "PARTS": [
                  `Part ${letterPart}`
            ]
        }
    };
    console.log(stresses);
    try {
      beamStresses = await midasAPI("POST", "/post/table", stresses);
      // setBeamStresses(beamStresses);
      console.log(beamStresses);
    } catch (error) {
      console.error("Error fetching beam stresses:", error);
    }
  
      for (let wkey in SelectWorksheets) {
        updatedata(wkey, SelectWorksheets[wkey]);
      }
      for (let wkey in SelectWorksheets2) {
        updatedata2(wkey, SelectWorksheets2[wkey], beamStresses);
      }
      for (let wkey in SelectWorksheets3) {
          updatedata3(wkey, SelectWorksheets3[wkey]);
        }
      if (!workbookData) return;
      const worksheet = workbookData.getWorksheet(sheetName);
      const buffer = await workbookData.xlsx.writeBuffer();
      const blob = new Blob([buffer], {
        type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
      });
  
      const link = document.createElement("a");
      link.href = URL.createObjectURL(blob);
      link.download = "output.xlsx";
      document.body.appendChild(link);
      link.click();
      document.body.removeChild(link);
    };
   
  
    const handleDataChange = (rowIndex, colIndex, value) => {
      const newData = [...sheetData];
      newData[rowIndex][colIndex] = value;
      setSheetData(newData);
    };
    function alert() {
      setCheck(true);
    }
  
    return (
      <Panel width={520} height={420} marginTop={3} padding={2} variant="shadow2">
        <div>
          <Typography variant="h1"> Casting Method</Typography>
          <RadioGroup
            margin={1}
            onChange={(e) => setCast(e.target.value)} // Update state variable based on user selection
            value={cast} // Bind the state variable to the RadioGroup
            text=""
          >
            <div
              style={{
                display: "flex",
                flexDirection: "row",
                alignItems: "start",
                justifyContent: "space-between",
                marginTop: "6px",
                marginRight: "5px",
                height: "10px",
                width: "198px",
              }}
            >
              <Radio
                name=" Cast In-Place"
                value="inplace"
                checked={cast === "inplace"}
              />
              <Radio
                name="Precast"
                value="precast"
                checked={cast === "precast"}
              />
            </div>
          </RadioGroup>
        </div>
  
        <div style={{ marginTop: "25px" }}>
          <Grid container>
            <Grid item xs={9}>
              <Typography variant="h1">
                {" "}
                Maximum Spacing of Transverse Reinforcement:
              </Typography>
            </Grid>
            <Grid item xs={3}>
              <Typography variant="h1"> (5.7.2.6.-1)</Typography>
            </Grid>
          </Grid>
          <RadioGroup
            margin={1}
            onChange={(e) => setSp(e.target.value)} // Update state variable based on user selection
            value={cast} // Bind the state variable to the RadioGroup
            text=""
          >
            <div
              style={{
                display: "flex",
                flexDirection: "row",
                alignItems: "start",
                justifyContent: "space-between",
                marginTop: "6px",
                marginRight: "5px",
                height: "10px",
                width: "300px",
              }}
            >
              <Radio name="CA (18 inches)" value="ca1" checked={sp === "ca1"} />
              <Radio
                name="AASHTO LFRD (24 inches)"
                value="aa1"
                checked={sp === "aa1"}
              />
            </div>
          </RadioGroup>
        </div>
  
        <div style={{ marginTop: "25px" }}>
          <Grid container>
            <Grid item xs={9}>
              <Typography variant="h1"> Clear Concrete Cover:</Typography>
            </Grid>
            <Grid item xs={3}>
              <Typography variant="h1"> (5.6.7-1)</Typography>
            </Grid>
          </Grid>
          <RadioGroup
            margin={1}
            onChange={(e) => setCvr(e.target.value)} // Update state variable based on user selection
            value={cast} // Bind the state variable to the RadioGroup
            text=""
          >
            <div
              style={{
                display: "flex",
                flexDirection: "row",
                alignItems: "start",
                justifyContent: "space-between",
                marginTop: "6px",
                marginRight: "5px",
                height: "10px",
                width: "300px",
              }}
            >
              <Radio name="CA (2.5 inches)" value="ca2" checked={cvr === "ca2"} />
              <Radio
                name="AASHTO LFRD (1.8 inches)"
                value="aa2"
                checked={cvr === "aa2"}
              />
            </div>
          </RadioGroup>
        </div>
        <div style={{ marginTop: "25px" }}>
          <Grid container>
            <Grid item xs={3}>
              <Typography variant="h1">
                {" "}
                Load Case for SLS (Permanent Loads)
              </Typography>
            </Grid>
            <Grid item xs={6} paddingLeft="10px">
              <DropList
                itemList={item}
                width="200px"
                defaultValue="Korean"
                value={value}
                onChange={onChangeHandler}
              />
            </Grid>
            <Grid item xs={3}>
              <Typography variant="h1">(5.9.2.3.2b-1)</Typography>
            </Grid>
          </Grid>
        </div>
  
        <div
          style={{
            display: "flex",
            flexDirection: "column",
            justifyContent: "space-between",
            marginTop: "20px",
          }}
        >
          <Grid container>
            <Grid item xs={6}>
              <Typography variant="h1" height="40px" paddingTop="15px">
                <input
                  type="file"
                  accept=".xlsx, .xls"
                  onChange={handleFileUpload}
                />
              </Typography>
            </Grid>
            <Grid item xs={6}>
              {/* <div
                              style={{
                                  borderBottom: "1px solid gray",
                                  height: "40px",
                                  width: "200px",
                                  display: "flex",
                                  justifyContent: "center",
                                  alignItems: "center",
                              }}
                          >
                              <div style={{ fontSize: "12px", paddingBottom: "2px" }}>                                
                              </div>
                          </div> */}
            </Grid>
          </Grid>
          {/*  */}
          <Grid container direction="row">
            <Grid item xs={6}>
              <Typography>Maximum aggregate size(ag) (in inches)</Typography>
            </Grid>
            <Grid item xs={6}>
              <TextField
                value={ag}
                onChange={handleAgChange}
                placeholder=""
                //   title="Maximum aggregate size(ag)"
                width="100px"
              />
            </Grid>
            <Grid item xs={6} marginTop={0.5}>
              <Typography size="small">
                Maximum distance between the longitudinal reinforcement (in
                inches)
              </Typography>
            </Grid>
            <Grid item xs={6} marginTop={0.5}>
              <TextField
                value={sg}
                onChange={handleSgChange}
                placeholder=""
                //   title="title"
                width="100px"
              />
            </Grid>
          </Grid>
          {/* </Grid>
                  </Grid> */}
        </div>
        <div
          style={{
            display: "flex",
            justifyContent: "space-between",
            margin: "0px",
            marginTop: "20px",
            marginBottom: "30px",
          }}
        >
          {/* {Buttons.NormalButton("contained", "Import Report", () => importReport())} */}
          {/* {Buttons.MainButton("contained", "Update Report", () => updatedata())}  */}
          {Buttons.MainButton("contained", "Create Report", handleFileDownload)}
          {check && <AlertDialogModal />}
        </div>
      </Panel>
    );
  };
  