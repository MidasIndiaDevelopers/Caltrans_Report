import {DropList,Grid,Panel,Typography,VerifyUtil, } from "@midasit-dev/moaui";
  import { Radio, RadioGroup } from "@midasit-dev/moaui";
  import React, { useState, useEffect, useRef } from "react";
  import * as Buttons from "./Components/Buttons";
  import ExcelJS from "exceljs";
  import AlertDialogModal from "./AlertDialogModal";
  import { midasAPI } from "./Function/Common";
  import { enqueueSnackbar } from "notistack";
  import { ThetaBeta1 } from "./Function/ThetaBeta";
  import { ThetaBeta2 } from "./Function/ThetaBeta";
  import { TextField } from "@midasit-dev/moaui";
  import { saveAs } from "file-saver";
  import { closeSnackbar } from 'notistack'
  import { useSnackbar, SnackbarProvider } from "notistack";
  import Image from '../src/assets/longitudianl_rf.png';

  
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
    const fileInputRef = useRef(null);
    const [buttonText, setButtonText] = useState('Create Report');
    let lcname;
    let mu_pos;
    let mu_neg;
    let mr_old_pos;
    let mr_new_pos;
    let mr_old_neg;
    let mr_new_neg;
    let check_mr_old;
    let check_mr_new;
    let s_m;
    let s_n;
    let sm_old;
    let sm_new;
    let sn_old;
    let sn_new;
    let smax_old; let smax_new; let suse; let dc_old; let dc_new; let beta_m; let theta_m; let beta_n; let theta_n; let Av_f; let Avr_old; let Avr_new; let vu; let vr_old; let vr_new; 
    let beta_mo; let beta_no; let theta_no; let theta_mo; let vr_old_n; let vr_new_n; let phi_new_m; let phi_new_n; 
    const action = snackbarId => (
        <>
          <button style={{ backgroundColor: 'transparent', border: 'none',color: 'white', cursor: 'pointer' }} onClick={() => { closeSnackbar(snackbarId) }}>
            Dismiss
          </button>
        </>
      );
  
    console.log(lcname);
    useEffect(() => {
      fetchLc();
    }, []); 
  
 
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
          let summarySheet = workbook.addWorksheet('Summary');
          let newMatchedParts = [];
          // const summaryWorkbook = await fetchAndProcessExcelFile();
          // const summarySheet = summaryWorkbook.getWorksheet('Sheet1');
  
        if (!summarySheet) {
          // throw new Error("Sheet1 not found in Summary_Caltrans.xlsx");
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
              const addRichText = (cell) => {
                if (cell && cell.richText) {
                    // Your logic for handling richText
                    console.log('Processing cell with richText:', cell.richText);
                }
            };
    
            // Iterate through each row and cell in the worksheet
            worksheet.eachRow((row) => {
                row.eachCell((cell) => {
                    addRichText(cell);
                });
            });
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
  
          let startRowNumbers = [];
          let endRowNumbers = [];
  
          // Find all start and end rows based on values in the first cell
          worksheet.eachRow((row, rowNumber) => {
            if (row.getCell(1).value === "$$strm1") {
              startRowNumbers.push(rowNumber);
            } else if (row.getCell(1).value === "$$theta_max") {
              rowNumber = rowNumber + 1;
              endRowNumbers.push(rowNumber);
            }
          });
  
          if (startRowNumbers.length === 0 || endRowNumbers.length === 0) {
            throw new Error(
              "Could not find the start or end markers ($$strm1 or $$fpo)"
            );
          }
  
          // Ensure we have matching start and end markers
          if (startRowNumbers.length !== endRowNumbers.length) {
            throw new Error(
              "Mismatched number of start ($$strm1) and end ($$fpo) markers."
            );
          }
  
          // Process rows between each startRowNumber and endRowNumber pair
          for (let i = 0; i < startRowNumbers.length; i++) {
            let startRowNumber = startRowNumbers[i];
            let endRowNumber = endRowNumbers[i];
  
            for (
              let rowNumber = startRowNumber + 1;
              rowNumber < endRowNumber;
              rowNumber++
            ) {
              let row = worksheet.getRow(rowNumber);
              row.eachCell({ includeEmpty: true }, (cell, colNumber) => {
                if (!cell.value) {
                  let colLetter = indexToLetter(colNumber - 1);
                  let address = colLetter + rowNumber;
                  row.getCell(colNumber).value = "";
                  row.getCell(colNumber)._address = address;
                }
              });
            }
          }
  
          const lastRowNumber = worksheet2.rowCount;
          const newRowNumber = lastRowNumber + 1;
          const newRow = worksheet2.getRow(newRowNumber);
            worksheet2.mergeCells(newRowNumber, 1, newRowNumber + 1, 14);
            const mergedCell = worksheet2.getCell(newRowNumber, 1);
            mergedCell.value = "Tensile Stress Limits in Prestressed Concrete At Service Limit State after Losses : No tension case    (As per CA-5.9.2.2b-1)";
            mergedCell.font = { bold: true, size: 12 };
            mergedCell.alignment = { vertical: 'middle' };
            newRow.commit();
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
    function cot(angleInDegrees) {
        const angleInRadians = angleInDegrees * (Math.PI / 180);
        return 1 / Math.tan(angleInRadians);
    }
      
    console.log(matchedParts);
    const [ag, setAg] = useState("");
    const [sg, setSg] = useState("");
  
    const handleAgChange = (event) => {
        const value = event.target.value;
        setAg(value);
        if (Number(value) < 0) {
            enqueueSnackbar("Error: Value should always be greater than zero", {
              variant: "error",
              anchorOrigin: {
                vertical: "top",
                horizontal: "center",
              },
              action,
            });
            return;
          }
          if (isNaN(value)) {
            enqueueSnackbar("Error: Please input numeric values only", {
              variant: "error",
              anchorOrigin: {
                vertical: "top",
                horizontal: "center",
              },
              action,
            });
            return;
          }
    };
  
    const handleSgChange = (event) => {
        const value = event.target.value;
        setSg(value);  
       
        if (Number(value) <= 0) {
            enqueueSnackbar("Error: Value should always be greater than zero", {
              variant: "error",
              anchorOrigin: {
                vertical: "top",
                horizontal: "center",
              },
              action,
            });
            return;
          }
          if (isNaN(value)) {
            enqueueSnackbar("Error: Please input numeric values only", {
              variant: "error",
              anchorOrigin: {
                vertical: "top",
                horizontal: "center",
              },
              action,
            });
            return;
          }
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
      let s ;
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
      let vu_check;
      let vu_check_min;
      let Mn_i=0;
      let phi_i = 0;
      let Mr_i = 0;
      let M_i = 0;
      let N_i = 0;
      let pi_i = 0;
      let Vp_i = 0;
      let dv_i = 0;
      let Vu1_i = 0;
      let vu_i = 0;
      let sm_i = 0;
      let s_i = 0;
      let Avm_i = 0;
      let Av_i = 0;
      let vc_i = 0;
      let alpha_i = 0;
      let vp_i = 0;
      let Vc_i = 0;
      let a_i = 0;
      let A_i = 0;
      let e_i = 0;
      let sx_i = 0;
      let sxe_i = 0;
      let b_i =  0;
      let theta_i = 0;
      let beta_i = 0;
      let vcvp_i = 0;
      let vs_i = 0;
      let t_i = 0;
      let sum_i = 0;
      let vn_i = 0;
      let vr_i = 0;
      let check_i = 0;
      let bv;
      let h;
      let Mcr_i = 0;
      let bs_i = 0;
      let I ;
      let cm;
      let cp;
      let st1_i = 0;
      let chge_i = 0;
     
      for (let key1 in rows) {
        if (
          rows[key1]._cells[0] != undefined &&
          rows[key1]._cells[0]._value.model.value == "$$type"
        ) {
            let cell2 = rows[key1]._cells[2];
            let add2 = cell2._address;
            data = { ...data,[add2] : 'CALTRANS'}
          let cell17 = rows[key1]._cells[17];
          let add17 = cell17._address;
          let cell17Value = cell17.value !== undefined ? cell17.value : null;
          if (cell17Value === "Composite") {
            type = "Composite";
          } else if (cell17Value === undefined) {
            type = "Box";
          } else {
            type = "Box";
          }
        }
        console.log(type);
        if (
          rows[key1]._cells[0] != undefined &&
          rows[key1]._cells[0]._value.model.value == "$$H"
        ) {
          h = rows[key1]._cells[27]._value.model.value;
        }
        if (
          rows[key1]._cells[0] != undefined &&
          rows[key1]._cells[0]._value.model.value == "$$d-c"
        ) {
          dc_old = rows[key1]._cells[11].value;
        }
  
        if (
          rows[key1]._cells[0] != undefined &&
          rows[key1]._cells[0]._value.model.value == "$$Mn"
        ) {
          Mn_i = Mn_i + 1;
          if (Mn_i == 1) {
          let location = rows[key1]._cells[19]._value.model.address;
          let value = rows[key1]._cells[19]._value.model.value;
          value = parseFloat(value.toFixed(3));
          data = { ...data, [location]: value };
          mn = value;
          }
          if (Mn_i ==2) {
           let location = rows[key1]._cells[19]._value.model.address;
          let value = rows[key1]._cells[19]._value.model.value;
          value = parseFloat(value.toFixed(3));
          data = { ...data, [location]: value };
          mn_neg = value;
          }
        }
        if (
          rows[key1]._cells[0] != undefined &&
          rows[key1]._cells[0]._value.model.value == "$$Phi"
        ) {
          phi_i = phi_i + 1;
          if (phi_i == 1) {
          let location = rows[key1]._cells[5]._value.model.address;
          if (cast === "inplace") {
            phi_new_m = 0.95;
            data = { ...data, [location]: 0.95 };
            phi = 0.95;
            let equ = rows[key1]._cells[22]._value.model.address;
            let existingValue = rows[key1]._cells[22]._value.model.value;
            // Concatenate the existing value with the new string
            let concatenatedValue = '0.005 ≤εt  ' + '(As per CA- 5.5.4.2)';
            data = { ...data, [equ]: concatenatedValue };
          } else {
            data = { ...data, [location]: 1 };
            phi = 1;
          }
          console.log(phi);
        }
          if (phi_i == 2) {
            let location = rows[key1]._cells[5]._value.model.address;
            if (cast === "inplace") {
              phi_new_n = 0.95;
              data = { ...data, [location]: 0.95 };
              phi = 0.95;
              let equ = rows[key1]._cells[25]._value.model.address;
              data = { ...data, [equ] : '(As per CA- 5.5.4.2)'}
            } else {
              data = { ...data, [location]: 1 };
              phi = 1;
            }
            console.log(phi);
          }
        }
        if (
          rows[key1]._cells[0] != undefined &&
          rows[key1]._cells[0]._value.model.value == "$$Mr"
        ) {
          Mr_i = Mr_i + 1;
          if(Mr_i == 1) {
          let location = rows[key1]._cells[5]._value.model.address;
  
          let mu = rows[key1]._cells[17]._value.model.value;
          let value5 = rows[key1]._cells[5]._value.model.value; 
          let check_mr = rows[key1]._cells[29]._value.model.value;
          check_mr_old = check_mr;
          mu_pos = mu;
          mr_old_pos = rows[key1]._cells[5]._value.model.value;
          mr = value5*phi;
          mr = parseFloat(mr.toFixed(3));
          mr_new_pos = mr;
          data = { ...data, [location]: mr };
  
          // location of oK
          if (Math.abs(mr) < Math.abs(Number(mu))) {
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
         if(Mr_i == 2) {
          let location = rows[key1]._cells[5]._value.model.address;
  
          let mu = rows[key1]._cells[17]._value.model.value;
          let value5 = rows[key1]._cells[5]._value.model.value; 
          mu_neg = mu;
          mr_neg = rows[key1]._cells[5]._value.model.value;
          mr_old_neg = mr_neg;
          mr_neg = value5*phi;
          mr_neg = parseFloat(mr_neg.toFixed(3));
          mr_new_neg = mr_neg;
          data = { ...data, [location]: mr_neg };
  
          // location of oK
          if (Math.abs(mr) < Math.abs(Number(mu))) {
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
      }
        if (
          getSafeCell(rows[key1], 0) &&
          getSafeCell(rows[key1], 0)._value.model.value === "$$Mcr"
        ) {
          Mcr_i += 1;
        
          let cell5 = getSafeCell(rows[key1], 5);
          let cell13 = getSafeCell(rows[key1], 13);
          let cell29 = getSafeCell(rows[key1], 29);
          let cell21 = getSafeCell(rows[key1], 21);
        
          let add5 = cell5._address;
          let add13 = cell13._address;
          let add29 = cell29._address;
        
          if (Mcr_i === 1) {
            data = { ...data, [add5]: mr_new_pos };
        
            if (Math.abs(cell21._value.model.value) < Math.abs(mr_new_pos)) {
              data = { ...data, [add13]: '≥' };
              data = { ...data, [add29]: 'OK' };
            } else {
              data = { ...data, [add13]: '<' };
              data = { ...data, [add29]: 'NG' };
            }
          }
        
          if (Mcr_i === 2) {
            data = { ...data, [add5]: mr_new_neg };
        
            if (Math.abs(cell21._value.model.value) < Math.abs(mr_new_neg)) {
              data = { ...data, [add13]: '≥' };
              data = { ...data, [add29]: 'OK' };
            } else {
              data = { ...data, [add13]: '<' };
              data = { ...data, [add29]: 'NG' };
            }
          }
        }
        
        if (
          rows[key1]._cells[0] != undefined &&
          rows[key1]._cells[0]._value.model.value == "$$dv"
        ) {
          dv_i = dv_i + 1;
          if (dv_i == 1) {
          dv = rows[key1]._cells[4]._value.model.value;
         }
         if (dv_i == 2) {
            dv_min = rows[key1]._cells[4]._value.model.value;
           }
        }
        if (
          rows[key1]._cells[0] != undefined &&
          rows[key1]._cells[0]._value.model.value == "$$t"
        ) { 
          theta_i = theta_i + 1;
          if( theta_i == 1) {
           
          }
          if (theta_i == 2) {
            theta_no = rows[key1]._cells[13].value ;
          }
        
        }
  
        if (
          rows[key1]._cells[0] != undefined &&
          rows[key1]._cells[0]._value.model.value == "$$vu"
        ) {
           vu_i = vu_i + 1;
           if (vu_i == 1) {
              vu_check = rows[key1]._cells[11]._value.model.value;
           }
           if (vu_i == 2) {
              vu_check = rows[key1]._cells[11]._value.model.value;
           }
           
        }
        
      if (
        rows[key1]._cells[0] != undefined &&
        rows[key1]._cells[0]._value.model.value == "$$sm"
      ) {
        sm_i = sm_i + 1;
        if (sm_i == 1) {
        sm_old = rows[key1]._cells[13]._value.model.value;
        if (sp === "ca1" && vu_check == '<') {         
          let add1 = rows[key1]._cells[6]._value.model.address;
          data = { ...data, [add1]: "Min[0.8dv, 18.0(in.)]" };
          let add2 = rows[key1]._cells[13]._value.model.address;
          // let val=rows[key1]._cells[13]._value.model.value;
          if (0.8 * dv >= 18) {
            data = { ...data, [add2]: 18 };
            sm_new = 18;
          } else {
            data = { ...data, [add2]: 0.8 * dv };
            sm_new = 0.8 * dv;
          }
          let add27 = rows[key1]._cells[27]._value.model.address;
          data = { ...data,[add27] : '(As per CA-5.7.2.6-1)'}
        }
        
      }
      if (sm_i == 2) {
         sn_old = rows[key1]._cells[13]._value.model.value;
          if (sp === "ca1" && vu_check == '<') {    
            let add1 = rows[key1]._cells[6]._value.model.address;
            data = { ...data, [add1]: "Min[0.8dv, 18.0(in.)]" };
            let add2 = rows[key1]._cells[13]._value.model.address;
            // let val=rows[key1]._cells[13]._value.model.value;
            if (0.8 * dv >= 18) {
              data = { ...data, [add2]: 18 };
              sn_new = 18;
            } else {
              data = { ...data, [add2]: 0.8 * dv };
              sn_new = 0.8 * dv;
            }
            let add27 = rows[key1]._cells[27]._value.model.address;
            data = { ...data,[add27] : '(As per CA-5.7.2.6-1)'}
          }
         
        }
      }
  
      if (
        rows[key1]._cells[0] != undefined &&
        rows[key1]._cells[0]._value.model.value == "$$s"
      ) {
          s_i = s_i + 1;
        if (s_i == 1) {
        s_m = rows[key1]._cells[8].value;
        let cell8 = rows[key1]._cells[8];
        s_max = rows[key1]._cells[8].value;
        console.log(s_max);
        }
          if (s_i == 2) {
            s_n = rows[key1]._cells[8].value;
        let cell8 = rows[key1]._cells[8];
        s_min = rows[key1]._cells[8].value;
        console.log(s_min);
       }
  
      }
  
       
        if (
          rows[key1]._cells[0] != undefined &&
          rows[key1]._cells[0]._value.model.value == "$$sx"
        ) {
          K = rows[key1]._cells[15].value;
        }
        if (
          rows[key1]._cells[0] != undefined &&
          rows[key1]._cells[0]._value.model.value == "$$dc"
        ) {
            smax_old = rows[key1]._cells[11].value;
            suse = rows[key1]._cells[21].value;
            let val2_new;
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
                  column15Address = rows[nextKey1]._cells[15]._value.model.address;
                  column15Value = rows[nextKey1]._cells[15]._value.model.value;
                  column15Value_new = (1 + (2.5/(0.7*(h-2.5))));
                  column15Value_new = parseFloat(column15Value_new.toFixed(3));
                  console.log(column15Value_new);
                  storedValues[column15Address] = column15Value;
                  data = { ...data, [column15Address]: column15Value_new };
                }
                if (rows[nextKey1]._cells[0]._value.model.value == "$$d-c") {
                  // Store the value and address of cell in column 9 for $$dc row
                  // dc_old = rows[nextKey1]._cells[11].value;
                  dc_new = 2.5;
                  column9Value = rows[nextKey1]._cells[9]._value.model.value;
                  column9Address = rows[nextKey1]._cells[9]._value.model.address;
                  column9Value_new = column9Value - column9Value + 2.5;
                  console.log(column9Value_new);
                  storedValues[column9Address] = column9Value;
                  data = { ...data, [column9Address]: column9Value_new };
                  let add13 = rows[nextKey1]._cells[13]._value.model.address;
                  data = { ...data,[add13] : '(in)                                                                    (As per CA-5.6.7-1)'};
                  break;
                }
              }
              nextKey1++;
            }
            console.log(storedValues);
            val2_new = ((val2 + 3.6) * column15Value) / column15Value_new - 5;
            val2_new = parseFloat(val2_new.toFixed(3));
  
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
          if (cvr === "ca2") {
            smax_new = val2_new;
          }
          else {
            smax_new = smax_old;
          }
        }
        if (
          rows[key1]._cells[0] != undefined &&
          rows[key1]._cells[0]._value.model.value == "$$a"
        ) {
          let cell8 = rows[key1]._cells[8];
          a = cell8.value;
        }
        // if (
        //   rows[key1]._cells[0] != undefined &&
        //   rows[key1]._cells[0]._value.model.value == "$$a_min"
        // ) {
        //   let cell8 = rows[key1]._cells[8];
        //   a = cell8.value;
        // }
        console.log(a);
        if (
          rows[key1]._cells[0] != undefined &&
          rows[key1]._cells[0]._value.model.value == "$$fy"
        ) {
          if ( type == 'Composite') {
          fy = rows[key1]._cells[6].value;
          }
          else {
            fy = rows[key1]._cells[7].value;
          }
        }
        console.log(fy);
  
        // if (
        //   rows[key1]._cells[0] != undefined &&
        //   rows[key1]._cells[0]._value.model.value == "$$Avm"
        // ) {
        //   if (type == "Composite") {
        //     Avm = rows[key1]._cells[12]._value.model.value;
        //   } else {
        //     Avm = rows[key1]._cells[17]._value.model.value;
        //   }
        // }
        if (
          rows[key1]._cells[0] != undefined &&
          rows[key1]._cells[0]._value.model.value == "$$Av"
        ) {
          Av = rows[key1]._cells[16]._value.model.value;
          Av_f = Av;
          s = rows[key1]._cells[20]._value.model.value;
        }
        if (
          rows[key1]._cells[0] != undefined &&
          rows[key1]._cells[0]._value.model.value == "$$M"
        ) {
            M_i = M_i + 1;
            if (M_i == 1) {
          Mmax = rows[key1]._cells[15]._value.model.value;
            }
            if (M_i == 2) {
                Mmin = rows[key1]._cells[15]._value.model.value;
            }
        }
  
        if (
          rows[key1]._cells[0] != undefined &&
          rows[key1]._cells[0]._value.model.value == "$$Ag"
        ) {
          Ag = rows[key1]._cells[27]._value.model.value;
          if (type == 'Box') {
            Ag = rows[key1]._cells[14]._value.model.value;
            h= rows[key1]._cells[4]._value.model.value;
            St = rows[key1]._cells[24]._value.model.value;
          }
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
          if(type == 'Box') {
            Sb = rows[key1]._cells[24]._value.model.value;
          }
        }
        if (
          rows[key1]._cells[0] != undefined &&
          rows[key1]._cells[0]._value.model.value == "$$N"
        ) {
          N_i = N_i + 1;
          if ( N_i == 1) {
          Nmax = rows[key1]._cells[15]._value.model.value;
          }
          if ( N_i == 2) {
            Nmin = rows[key1]._cells[15]._value.model.value;
          }
        }
  
        if (
          rows[key1]._cells[0] != undefined &&
          rows[key1]._cells[0]._value.model.value == "$$E"
        ) {
          E = rows[key1]._cells[12]._value.model.value;
          fc = rows[key1]._cells[5]._value.model.value;
          if ( type == 'Box'){
            E = rows[key1]._cells[9]._value.model.value;
            fc = rows[key1]._cells[2]._value.model.value;
          }
        }
        if (
          rows[key1]._cells[0] != undefined &&
          rows[key1]._cells[0]._value.model.value == "$$Vu1"
        ) {
          Vu1_i = Vu1_i + 1;
          if (Vu1_i == 1) {
          if (type == "Composite") {
            Vu1 = rows[key1]._cells[10]._value.model.value;
          } else {
            Vu1 = rows[key1]._cells[13]._value.model.value;
          }
        }
        if (Vu1_i == 2) {
            if (type == "Composite") {
                Vu2 = rows[key1]._cells[10]._value.model.value;
              } else {
                Vu2 = rows[key1]._cells[13]._value.model.value;
              }
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
          rows[key1]._cells[0]._value.model.value == "$$Iy"
        ) {
          I = rows[key1]._cells[27]._value.model.value;
        }
        if (
          rows[key1]._cells[0] != undefined &&
          rows[key1]._cells[0]._value.model.value == "$$Cps"
        ) {
          cp = rows[key1]._cells[27]._value.model.value;
        }
        if (
          rows[key1]._cells[0] != undefined &&
          rows[key1]._cells[0]._value.model.value == "$$Cms"
        ) {
          cm = rows[key1]._cells[27]._value.model.value;
        }
        if (
          rows[key1]._cells[0] != undefined &&
          rows[key1]._cells[0]._value.model.value == "$$bv"
        ) {
          bv = rows[key1]._cells[4]._value.model.value;
        }
        if (
          rows[key1]._cells[0] != undefined &&
          rows[key1]._cells[0]._value.model.value == "$$alpha"
        ) {
          alpha_i = alpha_i + 1;
          if (alpha_i == 1){
          let cell8 = rows[key1]._cells[8];
          $$alpha = rows[key1]._cells[8].value;
          }
          if (alpha_i == 2) {
            let cell8 = rows[key1]._cells[8];
            $$alpha_min = rows[key1]._cells[8].value;
  
          }
        }
  
        if (
          rows[key1]._cells[0] != undefined &&
          rows[key1]._cells[0]._value.model.value == "$$Vp"
        ) {
          vp_i = vp_i + 1;
          if (vp_i == 1) {
          let cell19 = rows[key1]._cells[19];
          Vp = rows[key1]._cells[19].value;
          }
          if(vp_i == 2) {
            let cell19 = rows[key1]._cells[19];
            Vp_min = rows[key1]._cells[19].value;
          }
        }
  
      }
      console.log(Ag);
      console.log(Sb);
      console.log(St);
      console.log(E);
      console.log(fc);
      // console.log(Avm);
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
      let Ecm;
      let Etm;
      let Ecn;
      let Etn;
      if ( type == 'Box') {
      Ecm =((-1 * Number(Mmax)) / Number(St) + Number(Nmax) / Number(Ag)) /
        Number(E);
      Ecn =((-1 * Number(Mmin)) / Number(St) + Number(Nmin) / Number(Ag)) /
        Number(E);
      Etm =((1 * Number(Mmax)) / Number(Sb) + Number(Nmax) / Number(Ag)) /
        Number(E);
      Etn =((1 * Number(Mmin)) / Number(Sb) + Number(Nmin) / Number(Ag)) /
        Number(E);
      }
      else {
        Ecm = ((-1 * Number(Mmax) * Number(cp)) / Number(I) + Number(Nmax) / Number(Ag)) /
        Number(E);
      Ecn =((-1 * Number(Mmin) * Number(cp)) / Number(I) + Number(Nmin) / Number(Ag)) /
        Number(E);
      Etm =((1 * Number(Mmax) * Number(cm)) / Number(I) + Number(Nmax) / Number(Ag)) /
        Number(E);
      Etn =((1 * Number(Mmin) * Number(cm) ) / Number(I) + Number(Nmin) / Number(Ag)) /
        Number(E);
      }
      // console.log(Vu1,Vu2,fc)
      let a1 = Number(Vu1) / Number(fc);
      let a2 = Number(Vu2) / Number(fc);
      let Exm = (Ecm + Etm) / 2;
      let Exn = (Ecn + Etn) / 2;
      // console.log(a1,a2, Exm * 1000, Exn * 1000)
      let value1 = ThetaBeta1(Exm * 1000,a1);
      let value2 = ThetaBeta1(Exn * 1000,a2);
      // console.log(value1,value2);
      let theta1 = value1[0];
      let theta2 = value2[0];
      beta1 = value1[1];
      let beta2 = value2[1];
      Vc1 = Vc / beta1;
      Avm = (0.0316 * Math.sqrt(fc) * s * bv) / fy;
      console.log(Avm);
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
      let sxe;
      theta_i = 0;
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
  
      let strm1_i =0;
      let strm_i = 0;
      let fpo_i = 0;
      function getSafeCell(row, index) {
        try {
          if (row && row._cells && row._cells[index] && row._cells[index]._value && row._cells[index]._value.model) {
            return row._cells[index];
          }
        } catch (error) {
          console.error(`Error accessing cell at index ${index}:`, error);
        }
        return null;
      }
      
      // Processing rows
      for (let key1 in rows) {
        let row = rows[key1];
      
        if (getSafeCell(row, 0) && getSafeCell(row, 0)._value.model.value === "Design Condition") {
          let add1 = getSafeCell(row, 0)._value.model.address;
          data = { ...data, [add1]: 'Design Condition for Caltrans Amendment As per AASHTO LRFD Bridge Design' };
        }
      
        worksheet.eachRow((row, rowNumber) => {
            if (getSafeCell(row, 0) && getSafeCell(row, 0)._value.model.value === "$$strm") {
              strm_i += 1;
              let add1, add2, add12;
          
              // Unmerge cells in the current row if they are merged
              for (let col = 1; col <= 8; col++) {
                let cell = row.getCell(col);
                if (cell.isMerged) {
                  worksheet.unMergeCells(cell.address);
                }
              }
          
              // Merge cells from A to H in the current row
              worksheet.mergeCells(`B${rowNumber}:U${rowNumber}`);
          
              const mergedCell = row.getCell(2);
              const mergedCellValue =
                " 4) Calculation for β and θ                               (As per CA - 5.7.3.4)";
              mergedCell.value = mergedCellValue;
              data = { ...data, [`B${rowNumber}`]: mergedCellValue };
  
              if (strm_i === 1) {
                add1 = getSafeCell(row, 2)._value.model.address;
                add2 = getSafeCell(row, 8)._value.model.address;
                data = { ...data, [add1]: mergedCellValue };
                data = { ...data, [add2]: "" };
              }
          
              if (strm_i === 2) {
                add1 = getSafeCell(row, 2)._value.model.address;
                add2 = getSafeCell(row, 8)._value.model.address;
                data = { ...data, [add1]: mergedCellValue };
                data = { ...data, [add2]: "" };
              }
            }
          });
        if (getSafeCell(row, 0) && getSafeCell(row, 0)._value.model.value === "$$strm1") {
          strm1_i += 1;
        //   if (strm1_i == 1) {
        //   for (let i = 1; i <= 50; i++) {
        //     let cell = getSafeCell(row, i);
        //     if (cell) {
        //       cell._value.model = { value: " " };
        //       cell.style = { font: { underline: false } };
        //     } else {
        //       row._cells[i] = {
        //         _value: {
        //           model: {
        //             value: "dummy",
        //             address: indexToLetter(i) + (parseInt(key1) + 1),
        //           },
        //         },
        //         _address: indexToLetter(i) + (parseInt(key1) + 1),
        //         style: { font: { underline: false } },
        //       };
        //     }
        //   }
    
        //   let nextKey = parseInt(key1) + 1;
        //   while (rows[nextKey] && getSafeCell(rows[nextKey], 0) && getSafeCell(rows[nextKey], 0)._value.model.value !== "$$theta_max") {
        //     for (let i = 1; i <= 50; i++) {
        //       let cell = getSafeCell(rows[nextKey], i);
        //       if (cell) {
        //         cell._value.model = { value: " " };
        //         cell.style = { font: { underline: false } };
        //       } else {
        //         rows[nextKey]._cells[i] = {
        //           _value: {
        //             model: {
        //               value: "dummy",
        //               address: indexToLetter(i) + (nextKey + 1),
        //             },
        //           },
        //           _address: indexToLetter(i) + (nextKey + 1),
        //           style: { font: { underline: false } },

        //         };
        //       }
        //     }
        //     nextKey++;
        //   }
        // }
        if (strm1_i == 1) {
          for (let i = 1; i <= 50; i++) {
            let cell = getSafeCell(row, i);
            if (cell) {
              cell._value.model = { value: " " };
              cell.style = { font: { underline: false } };
            } else {
              row._cells[i] = {
                _value: {
                  model: {
                    value: "dummy",
                    address: indexToLetter(i) + (parseInt(key1) + 1),
                  },
                },
                _address: indexToLetter(i) + (parseInt(key1) + 1),
                style: { font: { underline: false } },
              };
            }
          }
      
          let nextKey = parseInt(key1) + 1;
          while (rows[nextKey] && getSafeCell(rows[nextKey], 0) && getSafeCell(rows[nextKey], 0)._value.model.value !== "$$theta_max") {
            for (let i = 1; i <= 50; i++) {
              let cell = getSafeCell(rows[nextKey], i);
              if (cell) {
                cell._value.model = { value: " " };
                cell.style = { font: { underline: false } };
              } else {
                rows[nextKey]._cells[i] = {
                  _value: {
                    model: {
                      value: "dummy",
                      address: indexToLetter(i) + (nextKey + 1),
                    },
                  },
                  _address: indexToLetter(i) + (nextKey + 1),
                  style: { font: { underline: false } },
                };
              }
            }
            nextKey++;
          }
      
          // Process the row containing $$theta_max
          if (rows[nextKey] && getSafeCell(rows[nextKey], 0) && getSafeCell(rows[nextKey], 0)._value.model.value === "$$theta_max") {
            for (let i = 1; i <= 50; i++) {
              let cell = getSafeCell(rows[nextKey], i);
              if (cell) {
                cell._value.model = { value: " " };
                cell.style = { font: { underline: false } };
              } else {
                rows[nextKey]._cells[i] = {
                  _value: {
                    model: {
                      value: "dummy",
                      address: indexToLetter(i) + (nextKey + 1),
                    },
                  },
                  _address: indexToLetter(i) + (nextKey + 1),
                  style: { font: { underline: false } },
                };
              }
            }
            nextKey++;
          }
      
          // Process one extra row after the $$theta_max row
          if (rows[nextKey]) {
            for (let i = 1; i <= 50; i++) {
              let cell = getSafeCell(rows[nextKey], i);
              if (cell) {
                cell._value.model = { value: " " };
                cell.style = { font: { underline: false } };
              } else {
                rows[nextKey]._cells[i] = {
                  _value: {
                    model: {
                      value: "dummy",
                      address: indexToLetter(i) + (nextKey + 1),
                    },
                  },
                  _address: indexToLetter(i) + (nextKey + 1),
                  style: { font: { underline: false } },
                };
              }
            }
          }
        }

      
        if (strm1_i === 2) {
          for (let i = 1; i <= 50; i++) {
            let cell = getSafeCell(row, i);
            if (cell) {
              cell._value.model = { value: " " };
              cell.style = { font: { underline: false } };
            } else {
              row._cells[i] = {
                _value: {
                  model: {
                    value: "dummy",
                    address: indexToLetter(i) + (parseInt(key1) + 1),
                  },
                },
                _address: indexToLetter(i) + (parseInt(key1) + 1),
                style: { font: { underline: false } },
              };
            }
          }
      
          let nextKey = parseInt(key1) + 1;
          while (rows[nextKey] && getSafeCell(rows[nextKey], 0) && getSafeCell(rows[nextKey], 0)._value.model.value !== "$$theta_max") {
            for (let i = 1; i <= 50; i++) {
              let cell = getSafeCell(rows[nextKey], i);
              if (cell) {
                cell._value.model = { value: " " };
                cell.style = { font: { underline: false } };
              } else {
                rows[nextKey]._cells[i] = {
                  _value: {
                    model: {
                      value: "dummy",
                      address: indexToLetter(i) + (nextKey + 1),
                    },
                  },
                  _address: indexToLetter(i) + (nextKey + 1),
                  style: { font: { underline: false } },
                };
              }
            }
            nextKey++;
          }
      
          // Process the row containing $$theta_max
          if (rows[nextKey] && getSafeCell(rows[nextKey], 0) && getSafeCell(rows[nextKey], 0)._value.model.value === "$$theta_max") {
            for (let i = 1; i <= 50; i++) {
              let cell = getSafeCell(rows[nextKey], i);
              if (cell) {
                cell._value.model = { value: " " };
                cell.style = { font: { underline: false } };
              } else {
                rows[nextKey]._cells[i] = {
                  _value: {
                    model: {
                      value: "dummy",
                      address: indexToLetter(i) + (nextKey + 1),
                    },
                  },
                  _address: indexToLetter(i) + (nextKey + 1),
                  style: { font: { underline: false } },
                };
              }
            }
            nextKey++;
          }
      
          // Process one extra row after the $$theta_max row
          if (rows[nextKey]) {
            for (let i = 1; i <= 50; i++) {
              let cell = getSafeCell(rows[nextKey], i);
              if (cell) {
                cell._value.model = { value: " " };
                cell.style = { font: { underline: false } };
              } else {
                rows[nextKey]._cells[i] = {
                  _value: {
                    model: {
                      value: "dummy",
                      address: indexToLetter(i) + (nextKey + 1),
                    },
                  },
                  _address: indexToLetter(i) + (nextKey + 1),
                  style: { font: { underline: false } },
                };
              }
            }
          }
        }
        }
  
        if (getSafeCell(rows[key1], 0) && getSafeCell(rows[key1], 0)._value.model.value === "$$A") {
          A_i = A_i + 1;
          const cellsToClear = Array.from({ length: 50 }, (_, i) => i + 1);

          // Clear the cell values
          cellsToClear.forEach((col) => {
            let cell = getSafeCell(rows[key1], col);
            if (cell) {
              cell.value = ""; // Clear the cell value
              cell.font = { ...cell.font, underline: 'none' };
            } else {
              console.error(
                `Error: Unable to determine address for rows[key1]._cells[${col}]`
              );
            }
          });

          if (A_i == 1) {
            console.log(rows[key1]._cells);

            let cell = getSafeCell(rows[key1], 4);
            if (cell) {
              data = { ...data, [cell._address]: "Av" };
            } else {
              console.error(
                "Error: Unable to determine address for rows[key1]._cells[4]"
              );
            }
            // worksheet.mergeCells(`G${key1}:H${key1}`);
            let cell2 = getSafeCell(rows[key1], 6);
            if (cell2) {
              const cellAddress = cell2._address;
              const rowNumber = parseInt(cellAddress.replace(/\D/g, ''), 10); // Extract the row number from the address
            
              // Unmerge cells G and H for the specific row if they are already merged
              try {
                worksheet.unMergeCells(`G${rowNumber}:H${rowNumber}`);
              } catch (error) {
                console.log(`Cells G${rowNumber}:H${rowNumber} were not merged.`);
              }
            
              // Merge cells G and H for the specific row
              worksheet.mergeCells(`G${rowNumber}:H${rowNumber}`);
            
              // Add "Aₘᵢₙ" to the merged cell G and set the alignment to center
              let mergedCell = worksheet.getCell(`G${rowNumber}`);
              mergedCell.value = "Aₘᵢₙ";
              mergedCell.alignment = { vertical: 'middle', horizontal: 'center' };
            
              data = { ...data, [`G${rowNumber}`]: "Aₘᵢₙ" };
            } else {
              console.error(
                "Error: Unable to determine address for rows[key1]._cells[6]"
              );
            }
            

            let cell3 = getSafeCell(rows[key1], 5);
            if (cell3) {
              let comparisonSymbol = Av >= Avm ? "≥" : "<";
              data = { ...data, [cell3._address]: comparisonSymbol };
            } else {
              console.error(
                "Error: Unable to determine address for rows[key1]._cells[5]"
              );
            }
            
          }
          if (A_i == 2) {
            console.log(rows[key1]._cells);

            let cell = getSafeCell(rows[key1], 4);
            if (cell) {
              data = { ...data, [cell._address]: "Av" };
            } else {
              console.error(
                "Error: Unable to determine address for rows[key1]._cells[4]"
              );
            }

            let cell2 = getSafeCell(rows[key1], 6);
            if (cell2) {
              const cellAddress = cell2._address;
              const rowNumber = parseInt(cellAddress.replace(/\D/g, ''), 10); // Extract the row number from the address
            
              // Unmerge cells G and H for the specific row if they are already merged
              try {
                worksheet.unMergeCells(`G${rowNumber}:H${rowNumber}`);
              } catch (error) {
                console.log(`Cells G${rowNumber}:H${rowNumber} were not merged.`);
              }
            
              // Merge cells G and H for the specific row
              worksheet.mergeCells(`G${rowNumber}:H${rowNumber}`);
            
              // Add "Aₘᵢₙ" to the merged cell G and set the alignment to center
              let mergedCell = worksheet.getCell(`G${rowNumber}`);
              mergedCell.value = "Aₘᵢₙ";
              mergedCell.alignment = { vertical: 'middle', horizontal: 'center' };
            
              data = { ...data, [`G${rowNumber}`]: "Aₘᵢₙ" };
            } else {
              console.error(
                "Error: Unable to determine address for rows[key1]._cells[6]"
              );
            }
            

            let cell3 = getSafeCell(rows[key1], 5);
            if (cell3) {
              let comparisonSymbol = Av >= Avm ? "≥" : "<";
              data = { ...data, [cell3._address]: comparisonSymbol };
            } else {
              console.error(
                "Error: Unable to determine address for rows[key1]._cells[5]"
              );
            }
          }
        }
      
        if (getSafeCell(row, 0) && getSafeCell(row, 0)._value.model.value === "$$pi") {
          pi_i = pi_i + 1;
          if (pi_i == 1){
          pi = getSafeCell(row, 15) ? getSafeCell(row, 15)._value : null;
          }
          if (pi_i == 2){
            pi_min = getSafeCell(row, 15) ? getSafeCell(row, 15)._value : null;
          }
        }
      
        // if (getSafeCell(row, 0) && getSafeCell(row, 0)._value.model.value === "$$pi_min") {
        //   pi_min = getSafeCell(row, 15) ? getSafeCell(row, 15)._value : null;
        // }
      
        if (getSafeCell(row, 0) && getSafeCell(row, 0)._value.model.value === "$$e") {
          e_i = e_i + 1;
          if (e_i == 1){
          let cell = getSafeCell(row, 4);
          if (cell) {
            data = { ...data, [cell._address]: "εx" };
          }
          let cell2 = getSafeCell(row, 5);
          if (cell2) {
            data = { ...data, [cell2._address]: "=" };
          }
          let cell3 = getSafeCell(row, 6);
          if (cell3) {
            let cell3Value = Av >= Avm ? Exm : Etm;
            data = { ...data, [cell3._address]: cell3Value.toFixed(6) };
          }
        }
        if (e_i == 2){
          console.log(rows[key1]._cells);
  
          let cell = getSafeCell(rows[key1], 4);
          if (cell) {
            data = { ...data, [cell._address]: "εx" };
          } else {
            console.error("Error: Unable to determine address for rows[key1]._cells[4]");
          }
  
          let cell2 = getSafeCell(rows[key1], 5);
          if (cell2) {
            data = { ...data, [cell2._address]: "=" };
          } else {
            console.error("Error: Unable to determine address for rows[key1]._cells[5]");
          }
  
          let cell3 = getSafeCell(rows[key1], 6);
          if (cell3) {
            let cell3Value = Av >= Avm ? Exn : Etn;
            data = { ...data, [cell3._address]: cell3Value.toFixed(6) };
          }
        }
        }
        
  
        if (getSafeCell(rows[key1], 0) && getSafeCell(rows[key1], 0)._value.model.value === "$$sx") {
          sx_i = sx_i + 1;
          if (sx_i === 1) {
            rows[key1]._cells = rows[key1]._cells.map((cell) =>
              cell === "" ? undefined : cell
            );
            console.log(rows[key1]._cells);
        
            let cell = getSafeCell(rows[key1], 4);
            if (cell) {
              data = { ...data, [cell._address]: "sx" };
            } else {
              console.error("Error: Unable to determine address for rows[key1]._cells[4]");
            }
        
            let cell2 = getSafeCell(rows[key1], 5);
            if (cell2) {
              data = { ...data, [cell2._address]: "=" };
            } else {
              console.error("Error: Unable to determine address for rows[key1]._cells[5]");
            }
        
            let startCol = 7;
            let endCol = 22;
            let rowNumber = getSafeCell(rows[key1], startCol).row;
            try {
              let mergeRange = worksheet.getCell(
                `${worksheet.getColumn(startCol).letter}${rowNumber}:${worksheet.getColumn(endCol).letter}${rowNumber}`
              );
              if (!mergeRange.isMerged) {
                worksheet.mergeCells(rowNumber, startCol, rowNumber, endCol);
              }
            } catch (error) {
              console.error("Error merging cells: ", error);
            }
        
            let cell3 = worksheet.getCell(rowNumber, startCol);
            if (cell3) {
              let add13 = cell3._address;
              if (Av < Avm) {
                data = {
                  ...data,
                  [add13]: `Min| dv, maximum distance between the longitudinal r/f |`,
                };
        
                let cell4 = getSafeCell(rows[key1], 23);
                if (cell4) {
                  data = { ...data, [cell4._address]: "=" };
                } else {
                  console.error("Error: Unable to determine address for rows[key1]._cells[18]");
                }
        
                let cell5 = getSafeCell(rows[key1], 24);
                if (cell5) {
                  let add15 = cell5._address;
                  let add15value = dv < sg ? dv : sg;
                  data = { ...data, [add15]: add15value };
                } else {
                  console.error("Error: Unable to determine address for rows[key1]._cells[19]");
                }
              } else {
                data = { ...data, [add13]: `Not Required` };
              }
            } else {
              console.error("Error: Unable to determine address for rows[key1]._cells[7]");
            }
          }
        
          if (sx_i === 2) {
            rows[key1]._cells = rows[key1]._cells.map((cell) =>
              cell === "" ? undefined : cell
            );
            console.log(rows[key1]._cells);
        
            let cell = getSafeCell(rows[key1], 4);
            if (cell) {
              data = { ...data, [cell._address]: "sx" };
            } else {
              console.error("Error: Unable to determine address for rows[key1]._cells[4]");
            }
        
            let cell2 = getSafeCell(rows[key1], 5);
            if (cell2) {
              data = { ...data, [cell2._address]: "=" };
            } else {
              console.error("Error: Unable to determine address for rows[key1]._cells[5]");
            }
        
            let startCol = 7;
            let endCol = 22;
            let rowNumber = getSafeCell(rows[key1], startCol).row;
            try {
              let mergeRange = worksheet.getCell(
                `${worksheet.getColumn(startCol).letter}${rowNumber}:${worksheet.getColumn(endCol).letter}${rowNumber}`
              );
              if (!mergeRange.isMerged) {
                worksheet.mergeCells(rowNumber, startCol, rowNumber, endCol);
              }
            } catch (error) {
              console.error("Error merging cells: ", error);
            }
        
            let cell3 = worksheet.getCell(rowNumber, startCol);
            if (cell3) {
              let add13 = cell3._address;
              if (Av < Avm) {
                data = {
                  ...data,
                  [add13]: `Min| dv, maximum distance between the longitudinal r/f |`,
                };
        
                let cell4 = getSafeCell(rows[key1], 23);
                if (cell4) {
                  data = { ...data, [cell4._address]: "=" };
                } else {
                  console.error("Error: Unable to determine address for rows[key1]._cells[18]");
                }
        
                let cell5 = getSafeCell(rows[key1], 24);
                if (cell5) {
                  let add15 = cell5._address;
                  let add15value = dv < sg ? dv : sg;
                  data = { ...data, [add15]: add15value };
                } else {
                  console.error("Error: Unable to determine address for rows[key1]._cells[19]");
                }
              } else {
                data = { ...data, [add13]: `Not Required` };
              }
            } else {
              console.error("Error: Unable to determine address for rows[key1]._cells[7]");
            }
          }
        }
        
        if (getSafeCell(rows[key1], 0) && getSafeCell(rows[key1], 0)._value.model.value === "$$sxe") {
          sxe_i = sxe_i + 1;
          
          let cell5 = getSafeCell(rows[key1], 19);
          if (cell5) {
            let add15 = cell5._address;
            let add15value = dv < sg ? dv : sg;
            // data = { ...data, [add15]: add15value };
            sxe = (add15value * 1.38) / (ag + 0.63); 
          } else {
            console.error("Error: Unable to determine address for rows[key1]._cells[19]");
          }
          if (sxe_i === 1) {
            let cell3 = getSafeCell(rows[key1], 4);
            if (cell3) {
              data = { ...data, [cell3._address]: "sₓₑ" };
            }
        
            let cell4 = getSafeCell(rows[key1], 5);
            if (cell4) {
              data = { ...data, [cell4._address]: "=" };
            }
        
            let mergeStartCol = 7;
            let mergeEndCol = 16;
            let mergeRowNumber = getSafeCell(rows[key1], mergeStartCol).row;
            try {
              let mergeRange = worksheet.getCell(
                `${worksheet.getColumn(mergeStartCol).letter}${mergeRowNumber}:${worksheet.getColumn(mergeEndCol).letter}${mergeRowNumber}`
              );
              if (!mergeRange.isMerged) {
                worksheet.mergeCells(mergeRowNumber, mergeStartCol, mergeRowNumber, mergeEndCol);
              }
            } catch (error) {
              console.error("Error merging cells: ", error);
            }
        
            let cell5 = worksheet.getCell(mergeRowNumber, mergeStartCol);
            if (cell5) {
              let add15 = cell5._address;
              if (Av < Avm) {
                data = { ...data, [add15]: sxe };
              } else {
                data = { ...data, [add15]: "Not required" };
              }
            } else {
              console.error("Error: Unable to determine address for rows[key1]._cells[7]");
            }
          }
        
          if (sxe_i === 2) {
            let cell3 = getSafeCell(rows[key1], 4);
            if (cell3) {
              data = { ...data, [cell3._address]: "sₓₑ" };
            }
        
            let cell4 = getSafeCell(rows[key1], 5);
            if (cell4) {
              data = { ...data, [cell4._address]: "=" };
            }
        
            let mergeStartCol = 7;
            let mergeEndCol = 16;
            let mergeRowNumber = getSafeCell(rows[key1], mergeStartCol).row;
            try {
              let mergeRange = worksheet.getCell(
                `${worksheet.getColumn(mergeStartCol).letter}${mergeRowNumber}:${worksheet.getColumn(mergeEndCol).letter}${mergeRowNumber}`
              );
              if (!mergeRange.isMerged) {
                worksheet.mergeCells(mergeRowNumber, mergeStartCol, mergeRowNumber, mergeEndCol);
              }
            } catch (error) {
              console.error("Error merging cells: ", error);
            }
        
            let cell5 = worksheet.getCell(mergeRowNumber, mergeStartCol);
            if (cell5) {
              let add15 = cell5._address;
              if (Av < Avm) {
                data = { ...data, [add15]: sxe };
              } else {
                data = { ...data, [add15]: "Not required" };
              }
            } else {
              console.error("Error: Unable to determine address for rows[key1]._cells[7]");
            }
          }
        }
        
        if (getSafeCell(rows[key1], 0) && getSafeCell(rows[key1], 0)._value.model.value === "$$b") {
          b_i = b_i + 1;
          if (b_i === 1) {
            let cell5 = getSafeCell(rows[key1], 4);
            if (cell5) {
              data = { ...data, [cell5._address]: "β" };
            } else {
              console.error("Error: Unable to determine address for rows[key1]._cells[4]");
            }
        
            let cell6 = getSafeCell(rows[key1], 5);
            if (cell6) {
              data = { ...data, [cell6._address]: "=" };
            } else {
              console.error("Error: Unable to determine address for rows[key1]._cells[5]");
            }
        
            let cell7 = getSafeCell(rows[key1], 6);
            if (cell7) {
              if (Av < Avm) {
                let b_value = ThetaBeta2(Etm * 1000,sxe);
                let beta = b_value[0];
                console.log(beta);
                data = { ...data, [cell7._address]: parseFloat(beta.toFixed(2)) };
                beta_new = beta;
                beta_m = beta_new;
              } else {
                data = { ...data, [cell7._address]: parseFloat(beta1.toFixed(2)) };
                beta_new = parseFloat(beta1.toFixed(2));
                beta_m = beta_new;
              }
            } else {
              console.error("Error: Unable to determine address for rows[key1]._cells[6]");
            }
            let cell11 = getSafeCell(rows[key1], 11);
            if (cell11) {
              data = { ...data, [cell11._address]: "θ" };
            } else {
              console.error("Error: Unable to determine address for rows[key1]._cells[4]");
            }
        
            let cell12 = getSafeCell(rows[key1], 12);
            if (cell12) {
              data = { ...data, [cell12._address]: "=" };
            } else {
              console.error("Error: Unable to determine address for rows[key1]._cells[5]");
            }
            let cell13 = getSafeCell(rows[key1], 13);
                  if (cell13) {
                    if (Av < Avm) {
                      let theta_value = ThetaBeta2(Etn * 1000,sxe);
                      let theta = parseFloat(theta_value[1].toFixed(2));
                      console.log(theta);
                      data = { ...data, [cell13._address]: theta };
                      theta_new = theta;
                      theta_m = theta_new;
                    } else {
                      data = { ...data, [cell13._address]: parseFloat(theta2.toFixed(2)) };
                      theta_new = parseFloat(theta2.toFixed(2));
                      theta_m = theta_new;
                    }
                  } else {
                    console.error("Error: Unable to determine address for rows[key1]._cells[6]");
                  }
          }
        
          if (b_i === 2) {
            let cell5 = getSafeCell(rows[key1], 4);
            if (cell5) {
              data = { ...data, [cell5._address]: "β" };
            } else {
              console.error("Error: Unable to determine address for rows[key1]._cells[4]");
            }
        
            let cell6 = getSafeCell(rows[key1], 5);
            if (cell6) {
              data = { ...data, [cell6._address]: "=" };
            } else {
              console.error("Error: Unable to determine address for rows[key1]._cells[5]");
            }
        
            let cell7 = getSafeCell(rows[key1], 6);
            if (cell7) {
              if (Av < Avm) {
                let b_value = ThetaBeta2( Etn * 1000,sxe);
                let beta = parseFloat(b_value[0].toFixed(2));
                console.log(beta);
                data = { ...data, [cell7._address]: parseFloat(beta.toFixed(2)) };
                beta_new_min = beta;
                beta_n = beta_new_min;
              } else {
                data = { ...data, [cell7._address]: parseFloat(beta2.toFixed(2)) };
                beta_new_min = parseFloat(beta2.toFixed(2));
                beta_n = beta_new_min;
              }
            } else {
              console.error("Error: Unable to determine address for rows[key1]._cells[6]");
            }
            let cell11 = getSafeCell(rows[key1], 11);
            if (cell11) {
              data = { ...data, [cell11._address]: "θ" };
            } else {
              console.error("Error: Unable to determine address for rows[key1]._cells[4]");
            }
        
            let cell12 = getSafeCell(rows[key1], 12);
            if (cell12) {
              data = { ...data, [cell12._address]: "=" };
            } else {
              console.error("Error: Unable to determine address for rows[key1]._cells[5]");
            }
            let cell13 = getSafeCell(rows[key1], 13);
                  if (cell13) {
                    if (Av < Avm) {
                      let theta_value = ThetaBeta2(Etn * 1000,sxe);
                      let theta = parseFloat(theta_value[1].toFixed(2));
                      console.log(theta);
                      data = { ...data, [cell13._address]: theta };
                      theta_new_min = theta;
                      theta_n = theta_new_min;
                    } else {
                      data = { ...data, [cell13._address]: parseFloat(theta2.toFixed(2)) };
                      theta_new_min = parseFloat(theta2.toFixed(2));
                      theta_n = theta_new_min;
                    }
                  } else {
                    console.error("Error: Unable to determine address for rows[key1]._cells[6]");
                  }
          }
        }
        
        if (getSafeCell(rows[key1], 0) && getSafeCell(rows[key1], 0)._value.model.value === "$$theta_max") {
          theta_i = theta_i + 1;
          
          if (theta_i === 1) {
        
            for (let i = 2; i <= 32; i++) {
              let cell = getSafeCell(rows[key1], i);
              if (cell) {
                data = { ...data, [cell._address]: "" };
              } else {
                console.error(`Error: Unable to determine address for rows[key1]._cells[${i}]`);
              }
            }
          }
          if (theta_i === 2) {
        
            for (let i = 2; i <= 32; i++) {
              let cell = getSafeCell(rows[key1], i);
              if (cell) {
                data = { ...data, [cell._address]: "" };
              } else {
                console.error(`Error: Unable to determine address for rows[key1]._cells[${i}]`);
              }
            }
          }
        }
        if (
          getSafeCell(rows[key1], 0) &&
          getSafeCell(rows[key1], 0)._value.model.value === "$$chge"
      ) { 
          chge_i = chge_i + 1;
          if (chge_i == 1) {
           let cell2 = rows[key1]._cells[2];
           let add2 = cell2._address;
           data = { ...data, [add2] : '0.5Φ(Vc+Vp)'};
           let cell11 = rows[key1]._cells[11];
           let add11 = cell11._address;
           data = { ...data, [add11] : 'Φ(Vc+Vp)'};
          }
          if (chge_i == 2) {
            let cell2 = rows[key1]._cells[2];
            let add2 = cell2._address;
            data = { ...data, [add2] : '0.5Φ(Vc+Vp)'};
            let cell11 = rows[key1]._cells[11];
            let add11 = cell11._address;
            data = { ...data, [add11] : 'Φ(Vc+Vp)'};
           }
      }
        
        if (getSafeCell(rows[key1], 0) && getSafeCell(rows[key1], 0)._value.model.value === "$$beta") {
          beta_i = beta_i + 1;
          if (beta_i === 1) {
            let cell9 = getSafeCell(rows[key1], 9);
            if (cell9) {
              beta = cell9._value.model.value;
              beta_mo = beta;
            } else {
              console.error("Error: Unable to determine address for rows[key1]._cells[9]");
            }
          }
        
          if (beta_i === 2) {
            let cell9 = getSafeCell(rows[key1], 9);
            if (cell9) {
              beta_min = cell9._value.model.value;
              beta_no = beta_min;
            } else {
              console.error("Error: Unable to determine address for rows[key1]._cells[9]");
            }
          }
        }
        
      if (getSafeCell(rows[key1], 0) && getSafeCell(rows[key1], 0)._value.model.value === "$$vc") {
        vc_i = vc_i + 1;
        if (vc_i == 1) {
        let cell13 = getSafeCell(rows[key1], 13);
        let cell9 = getSafeCell(rows[key1], 9);
        let cell5 = getSafeCell(rows[key1], 5);
      
        if (cell5 && cell9 && cell13) {
          let add5 = cell5._address;
          let add9 = cell9._address;
          data = { ...data, [add5]: "0.0316" };
      
          console.log(cell9._value.model.value);
          console.log(rows[key1]);
      
          if (cell13._address) {
            let add13 = cell13._address;
            let initialValue13 = cell13._value.model.value;
      
            let result;
            if (type === "Composite") {
              result = (initialValue13 / beta) * beta1;
            } else {
              result = (initialValue13 / K) * beta1;
            }
      
            cell13._value.model.value = result;
            data = { ...data, [add9]: "β √f'c bvdv" };
            data = { ...data, [add13]: parseFloat(result.toFixed(2)) };
      
            Vc = result; // new Vc value
          } else {
            console.error("Error: Unable to determine address for cell13");
          }
        } else {
          console.error("Error: Unable to determine address for cells[5], [9], or [13]");
        }
      }
      if (vc_i == 2) {
        let cell13 = getSafeCell(rows[key1], 13);
        let cell9 = getSafeCell(rows[key1], 9);
        let cell5 = getSafeCell(rows[key1], 5);
      
        if (cell5 && cell9 && cell13) {
          let add5 = cell5._address;
          let add9 = cell9._address;
          data = { ...data, [add5]: "0.0316" };
      
          console.log(cell9._value.model.value);
          console.log(rows[key1]);
      
          if (cell13._address) {
            let add13 = cell13._address;
            let initialValue13 = cell13._value.model.value;
      
            let result;
            if (type === "Composite") {
              result = (initialValue13 / beta_min) * beta2;
            } else {
              result = (initialValue13 / K) * beta2;
            }
      
            cell13._value.model.value = result;
            data = { ...data, [add9]: "β √f'c bvdv" };
            data = { ...data, [add13]: parseFloat(result.toFixed(2)) };
      
            Vc_min = result; // new Vc_min value
          } else {
            console.error("Error: Unable to determine address for cell13");
          }
        } else {
          console.error("Error: Unable to determine address for cells[5], [9], or [13]");
        }
      }
      }
      
      console.log("Initial Value:", initialValue13);
      console.log("New Value:", Vc);
      if (getSafeCell(rows[key1], 0) && getSafeCell(rows[key1], 0)._value.model.value === "$$(vc+vp)") {
        vcvp_i += 1;
        
        if (vcvp_i === 1) {
          let cell11 = getSafeCell(rows[key1], 11);
          let cell2 = getSafeCell(rows[key1], 2);
      
          if (cell11 && cell11._value.model.value !== undefined) {
            Vu_max = getSafeCell(rows[key1], 20)._value.model.value;
            finalResult = pi * (Vc + Vp);
            half_finalResult = finalResult / 2;
            console.log(finalResult);
            data = { ...data, [cell11._address]: parseFloat(finalResult.toFixed(2)) };
            data = { ...data, [cell2._address]: parseFloat(half_finalResult.toFixed(2)) };
          } else {
            console.error("Error: Unable to retrieve value for rows[key1]._cells[11]");
          }
        }
        if (vcvp_i === 2) {
          let cell11 = getSafeCell(rows[key1], 11);
          let cell2 = getSafeCell(rows[key1], 2);
      
          if (cell11 && cell11._value.model.value !== undefined) {
            Vu_min = getSafeCell(rows[key1], 20)._value.model.value;
            finalResult = pi * (Vc_min + Vp_min);
            half_finalResult = finalResult / 2;
            console.log(finalResult);
            data = { ...data, [cell11._address]: parseFloat(finalResult.toFixed(2)) };
            data = { ...data, [cell2._address]: parseFloat(half_finalResult.toFixed(2)) };
          } else {
            console.error("Error: Unable to retrieve value for rows[key1]._cells[11]");
          }
        }
      }
      if (getSafeCell(rows[key1], 0) && getSafeCell(rows[key1], 0)._value.model.value === "$$check") {
          check_i += 1;
        
          if (check_i === 1) {
            let cell2 = getSafeCell(rows[key1], 2);
            let cell11 = getSafeCell(rows[key1], 11);
            let cell12 = getSafeCell(rows[key1], 12);
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
              cell2Value = "Vu ≥ 0.5Φ(Vc+Vp)";
              data = { ...data, [add2]: "Vu ≥ 0.5Φ(Vc+Vp)" };
            }
        
            let key2 = parseInt(key1) + 1;
        
            if (cell2Value === "Vu ≥ 0.5Φ(Vc+Vp)") {
              if (getSafeCell(rows[key2], 0) && getSafeCell(rows[key2], 0)._value.model.value === "$$Ar") {
                let cell13 = getSafeCell(rows[key2], 13);
                let add13 = cell13._address;
                let Av_extra;
                let Avr = ((Vu_max - finalResult) * s_max) / (pi * fy * dv * (cot(theta_new) + cot(a)) * Math.sin(a));
                Avr = parseFloat(Avr.toFixed(3));
                console.log(Avr);
                data = { ...data, [add13]: Avr };
        
                for (let i = key2; i <= worksheet.rowCount; i++) {
                  let nextRow = worksheet.getRow(i);
        
                  if (getSafeCell(rows[i], 0) && getSafeCell(rows[i], 0)._value.model.value === "$$Av,req") {
                    Avr_old = rows[key2]._cells[13].value;
                    let cell12 = getSafeCell(rows[i], 12);
                    let add12 = cell12._address;
        
                    if (Avm > Avr) {
                      Av_extra = Avm;
                      data = { ...data, [add12]: parseFloat(Avm.toFixed(3)) };
                      Avr_new = parseFloat(Avm.toFixed(3));
                    } else {
                      Av_extra = Avr;
                      data = { ...data, [add12]: parseFloat(Avr.toFixed(3)) };
                      Avr_new = parseFloat(Avr.toFixed(3));
                    }
                  }
        
                  if (getSafeCell(rows[i], 0) && getSafeCell(rows[i], 0)._value.model.value === "$$A,v") {
                    let cell11 = getSafeCell(rows[i], 11);
                    let cell29 = getSafeCell(rows[i], 29);
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
                    break;
                  }
                }
              }
               else {
                // let key3 = parseInt(key1) + 2;
                // key3 += 5;
        
                // let cell19 = getSafeCell(rows[key3], 19);
                // let add19 = cell19._address;
                // data = { ...data, [add19]: "Av,req1" };
                // let cell20 = getSafeCell(rows[key3], 20);
                // let add20 = cell20._address;
                // data = { ...data, [add20]: "=" };
                // let cell21 = getSafeCell(rows[key3], 21);
                // let add21 = cell21._address;
                // data = { ...data, [add21]: "{ Vu - Φ(Vc+Vp) }·s" };
                // let cell21_n = getSafeCell(rows[key3 + 1], 21);
                // let add21_n = cell21_n._address;
                // data = { ...data, [add21_n]: "Φ·fy·dv(cotθ+cotα)sinα" };
                // let cell27 = getSafeCell(rows[key3], 27);
                // let add27 = cell27._address;
                // data = { ...data, [add27]: "=" };
                // let cell28 = getSafeCell(rows[key3], 28);
                // let add28 = cell28._address;
                // let Av_extra;
                // let Avr = ((Vu_max - finalResult) * s_max) / (pi * fy * dv * (cot(theta_new) + cot(a)) * Math.sin(a));
                // data = { ...data, [add28]: Avr };
              }
               } else {
              if (getSafeCell(rows[key2], 0) && getSafeCell(rows[key2], 0)._value.model.value === "$$Ar") {
                for (let i = parseInt(key2) + 1; i <= worksheet.rowCount; i++) {
                  let nextRow = worksheet.getRow(i);
                  let cell1 = nextRow.getCell(1);
        
                  if (cell1 && cell1.value === "$$A,v") {
                    nextRow.eachCell({ includeEmpty: true }, (cell) => {
                      cell.value = "";
                    });
                    break;
                  }
        
                  nextRow.eachCell({ includeEmpty: true }, (cell) => {
                    cell.value = "";
                  });
                }
              }
            }
          }
        
          // if (check_i === 2) {
          //   let cell2 = getSafeCell(rows[key1], 2);
          //   let cell11 = getSafeCell(rows[key1], 11);
          //   let cell12 = getSafeCell(rows[key1], 12);
          //   let add2 = cell2._address;
          //   let add11 = cell11._address;
          //   let add12 = cell12._address;
          //   let cell2Value;
        
          //   if (Math.abs(half_finalResult) > Vu_min) {
          //     data = { ...data, [add2]: "Vu < 0.5Φ(Vc+Vp)" };
          //     data = { ...data, [add11]: "∴" };
          //     data = { ...data, [add12]: "No Shear reinforcing" };
          //   } else {
          //     data = { ...data, [add2]: "Vu ≥ 0.5Φ(Vc+Vp)" };
          //   }
        
          //   let key2 = parseInt(key1) + 1;
        
          //   if (cell2Value === "Vu ≥ 0.5Φ(Vc+Vp)") {
          //     if (getSafeCell(rows[key2], 0) && getSafeCell(rows[key2], 0)._value.model.value === "$$Ar") {
          //       for (let i = parseInt(key2) + 1; i <= worksheet.rowCount; i++) {
          //         let nextRow = worksheet.getRow(i);
          //         let cell1 = nextRow.getCell(1);
        
          //         if (cell1 && cell1.value === "$$A,v") {
          //           nextRow.eachCell({ includeEmpty: true }, (cell) => {
          //             cell.value = "";
          //           });
          //           break;
          //         }
        
          //         nextRow.eachCell({ includeEmpty: true }, (cell) => {
          //           cell.value = "";
          //         });
          //       }
          //     } else {
          //       let key3 = parseInt(key1) + 2;
        
          //       if (getSafeCell(rows[key3], 0) && getSafeCell(rows[key3], 0)._value.model.value === "$$vs") {
          //         let cell19 = getSafeCell(rows[key3], 19);
          //         let add19 = cell19._address;
          //         data = { ...data, [add19]: "Av,req1" };
          //         let cell20 = getSafeCell(rows[key3], 20);
          //         let add20 = cell20._address;
          //         data = { ...data, [add20]: "=" };
          //         let cell21 = getSafeCell(rows[key3], 21);
          //         let add21 = cell21._address;
          //         data = { ...data, [add21]: "{ Vu - Φ(Vc+Vp) }·s" };
          //         let cell21_n = getSafeCell(rows[key3 + 1], 21);
          //         let add21_n = cell21_n._address;
          //         data = { ...data, [add21_n]: "Φ·fy·dv(cotθ+cotα)sinα" };
          //         let cell27 = getSafeCell(rows[key3], 27);
          //         let add27 = cell27._address;
          //         data = { ...data, [add27]: "=" };
          //         let cell28 = getSafeCell(rows[key3], 28);
          //         let add28 = cell28._address;
          //         let Av_extra;
          //         let Avr = ((Vu_max - finalResult) * s_max) / (pi * fy * dv * (cot(theta_new) + cot(a)) * Math.sin(a));
          //         data = { ...data, [add28]: Avr };
          //       }
          //     }
          //   } else {
          //     if (rows[key2] && rows[key2]._cells && rows[key2]._cells[0] && rows[key2]._cells[0]._value.model.value === "$$Ar") {
          //       for (let i = key2 + 1; i <= worksheet.rowCount; i++) {
          //         let nextRow = worksheet.getRow(i);
        
          //         if (nextRow.getCell(1).value === "$$A,v") {
          //           nextRow.eachCell({ includeEmpty: true }, (cell) => {
          //             cell.value = "";
          //           });
          //           break;
          //         }
        
          //         nextRow.eachCell({ includeEmpty: true }, (cell) => {
          //           cell.value = "";
          //         });
          //       }
          //     }
          //   }
          // }
          if (check_i === 2) {
            let cell2 = getSafeCell(rows[key1], 2);
            let cell11 = getSafeCell(rows[key1], 11);
            let cell12 = getSafeCell(rows[key1], 12);
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
              cell2Value = "Vu ≥0.5Φ(Vc+Vp)";
              data = { ...data, [add2]: "Vu ≥ 0.5Φ(Vc+Vp)" };
            }
        
            let key2 = parseInt(key1) + 1;
        
            if (cell2Value === "0.5Φ(Vc+Vp)") {
              if (getSafeCell(rows[key2], 0) && getSafeCell(rows[key2], 0)._value.model.value === "$$Ar") {
                let cell13 = getSafeCell(rows[key2], 13);
                let add13 = cell13._address;
                let Av_extra;
                let Avr = ((Vu_max - finalResult) * s_max) / (pi * fy * dv * (cot(theta_new) + cot(a)) * Math.sin(a));
                Avr = parseFloat(Avr.toFixed(3));
                console.log(Avr);
                data = { ...data, [add13]: Avr };
        
                for (let i = key2; i <= worksheet.rowCount; i++) {
                  let nextRow = worksheet.getRow(i);
        
                  if (getSafeCell(rows[i], 0) && getSafeCell(rows[i], 0)._value.model.value === "$$Av,req") {
                    Avr_old = rows[key2]._cells[13].value;
                    let cell12 = getSafeCell(rows[i], 12);
                    let add12 = cell12._address;
        
                    if (Avm > Avr) {
                      Av_extra = Avm;
                      data = { ...data, [add12]: parseFloat(Avm.toFixed(3)) };
                      Avr_new = parseFloat(Avm.toFixed(3));
                    } else {
                      Av_extra = Avr;
                      data = { ...data, [add12]: parseFloat(Avr.toFixed(3)) };
                      Avr_new = parseFloat(Avr.toFixed(3));
                    }
                  }
        
                  if (getSafeCell(rows[i], 0) && getSafeCell(rows[i], 0)._value.model.value === "$$A,v") {
                    let cell11 = getSafeCell(rows[i], 11);
                    let cell29 = getSafeCell(rows[i], 29);
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
                    break;
                  }
                }
              }
               else {
                // let key3 = parseInt(key1) + 2;
                // key3 += 5;
        
                // let cell19 = getSafeCell(rows[key3], 19);
                // let add19 = cell19._address;
                // data = { ...data, [add19]: "Av,req1" };
                // let cell20 = getSafeCell(rows[key3], 20);
                // let add20 = cell20._address;
                // data = { ...data, [add20]: "=" };
                // let cell21 = getSafeCell(rows[key3], 21);
                // let add21 = cell21._address;
                // data = { ...data, [add21]: "{ Vu - Φ(Vc+Vp) }·s" };
                // let cell21_n = getSafeCell(rows[key3 + 1], 21);
                // let add21_n = cell21_n._address;
                // data = { ...data, [add21_n]: "Φ·fy·dv(cotθ+cotα)sinα" };
                // let cell27 = getSafeCell(rows[key3], 27);
                // let add27 = cell27._address;
                // data = { ...data, [add27]: "=" };
                // let cell28 = getSafeCell(rows[key3], 28);
                // let add28 = cell28._address;
                // let Av_extra;
                // let Avr = ((Vu_max - finalResult) * s_max) / (pi * fy * dv * (cot(theta_new) + cot(a)) * Math.sin(a));
                // data = { ...data, [add28]: Avr };
              }
               } else {
              if (getSafeCell(rows[key2], 0) && getSafeCell(rows[key2], 0)._value.model.value === "$$Ar") {
                for (let i = parseInt(key2) + 1; i <= worksheet.rowCount; i++) {
                  let nextRow = worksheet.getRow(i);
                  let cell1 = nextRow.getCell(1);
        
                  if (cell1 && cell1.value === "$$A,v") {
                    nextRow.eachCell({ includeEmpty: true }, (cell) => {
                      cell.value = "";
                    });
                    break;
                  }
        
                  nextRow.eachCell({ includeEmpty: true }, (cell) => {
                    cell.value = "";
                  });
                }
              }
            }
          }
        }
      if (
        getSafeCell(rows[key1], 0) &&
        getSafeCell(rows[key1], 0)._value.model.value === "$$vs"
    ) {
        vs_i = vs_i + 1;
        if (vs_i === 1) {
            let cell13 = getSafeCell(rows[key1], 13);
            let add13 = cell13 ? cell13._address : null;
            if (add13) {
                let cal = ((Av * fy * dv * (cot(theta_new) + cot($$alpha))) * Math.sin($$alpha)) / s_max;
                cal = parseFloat(cal.toFixed(3));
                data = { ...data, [add13]: cal };
                Vs = cal;
            } else {
                console.error("Error: Unable to retrieve address for rows[key1]._cells[13]");
            }
            let cell5 = rows[key1]._cells[5];
            let add5 = cell5._address;
            data = { ...data, [add5] : 'Av·fy·dv(cotθ+cotα)sinα'};
        }
        if (vs_i === 2) {
            let cell13 = getSafeCell(rows[key1], 13);
            let add13 = cell13 ? cell13._address : null;
            if (add13) {
                let cal = (Av * fy * dv_min * (cot(theta_new_min) + cot($$alpha_min)) * Math.sin($$alpha_min)) / s_min;
                cal = parseFloat(cal.toFixed(3));
                data = { ...data, [add13]: cal };
                Vs_min = cal;
            } else {
                console.error("Error: Unable to retrieve address for rows[key1]._cells[13]");
            }
            let cell5 = rows[key1]._cells[5];
            let add5 = cell5._address;
            data = { ...data, [add5] : 'Av·fy·dv(cotθ+cotα)sinα'};
        }
    }
    if (
      getSafeCell(rows[key1], 0) &&
      getSafeCell(rows[key1], 0)._value.model.value === "$$sum"
  ) {
      sum_i = sum_i + 1;
      if (sum_i === 1) {
          let cell3 = getSafeCell(rows[key1], 3);
          let add3 = cell3._address;
          data = { ...data,[add3] : 'Vc +Vs +Vp'}
          let cell14 = getSafeCell(rows[key1], 14);
          let add14 = cell14._address;
          data = { ...data,[add14] : "0.25·f'cbvdv + Vp ="};
          let cell7 = getSafeCell(rows[key1], 7);
          let add7 = cell7 ? cell7._address : null;
          if (add7) {
              let sum = Vp + Vc + Vs;
              Vn = sum;
              data = { ...data, [add7]: parseFloat(Vn.toFixed(2)) };
  
              let cell13 = getSafeCell(rows[key1], 13);
              let cell20 = getSafeCell(rows[key1], 20);
              let add13 = cell13 ? cell13._address : null;
              let add20 = cell20 ? cell20._address : null;
              let value20 = cell20 ? cell20.value : null;
              if( type == 'Box') {
                value20 = (0.25*fc*bv*dv) + Vp;
                data = { ...data,[add20] : parseFloat(value20.toFixed(3))};
              }
              if (add13 && value20 !== null) {
                  data = { ...data, [add13]: Vn < value20 ? "≤" : ">" };
              } else {
                  console.error("Error: Unable to retrieve address or value for rows[key1]._cells[20]");
              }
          } else {
              console.error("Error: Unable to retrieve address for rows[key1]._cells[7]");
          }
      }
  
      if (sum_i === 2) {
          let cell7 = getSafeCell(rows[key1], 7);
          let add7 = cell7 ? cell7._address : null;
          if (add7) {
              let sum = Vp_min + Vc_min + Vs_min;
              Vn_min = sum;
              data = { ...data, [add7]: parseFloat(Vn_min.toFixed(2)) };
  
              let cell13 = getSafeCell(rows[key1], 13);
              let cell20 = getSafeCell(rows[key1], 20);
              let add13 = cell13 ? cell13._address : null;
              let add20 = cell20 ? cell20._address : null;
              let value20 = cell20 ? cell20.value : null;
              if( type == 'Box') {
                value20 = (0.25*fc*bv*dv) + Vp;
                data = { ...data,[add20] : parseFloat(value20.toFixed(3))};
              }
  
              if (add13 && value20 !== null) {
                  data = { ...data, [add13]: Vn_min < value20 ? "≤" : ">" };
              } else {
                  console.error("Error: Unable to retrieve address or value for rows[key1]._cells[20]");
              }
          } else {
              console.error("Error: Unable to retrieve address for rows[key1]._cells[7]");
          }
      }
  }
  if (
    rows[key1]._cells[0] != undefined &&
    rows[key1]._cells[0]._value.model.value == "$$t"
  ) {
    t_i = t_i + 1;
    if(t_i == 1) {
      let cell8 = rows[key1]._cells[8];
      let add8 = cell8._address;
      data = { ...data , [add8] : "####"}
      theta_mo = rows[key1]._cells[13].value;
      let cell13 = rows[key1]._cells[13];
      let add13 = cell13._address;
      data = { ...data,[add13] : "####"}
      let cell27 = rows[key1]._cells[27];
      let add27 = getCellAddress(cell27, "dummy_address_27");
      data = { ...data,[add27] : "  "}
    }
    if(t_i == 2) {
      let cell8 = rows[key1]._cells[8];
      let add8 = cell8._address;
      data = { ...data , [add8] : "####"}
      theta_mo = rows[key1]._cells[13].value;
      let cell13 = rows[key1]._cells[13];
      let add13 = cell13._address;
      data = { ...data,[add13] : "####"}
      let cell27 = rows[key1]._cells[27];
      let add27 = getCellAddress(cell27, "dummy_address_27");
      data = { ...data,[add27] : "  "}
    }
    if(t_i == 3) {
      let cell8 = rows[key1]._cells[8];
      let add8 = cell8._address;
      data = { ...data , [add8] : "####"}
      let cell13 = rows[key1]._cells[13];
      let add13 = cell13._address;
      data = { ...data,[add13] : "####"}
      let cell27 = rows[key1]._cells[27];
      let add27 = getCellAddress(cell27, "dummy_address_27");
      data = { ...data,[add27] : "  "}
    }
    if(t_i == 4) {
      let cell8 = rows[key1]._cells[8];
      let add8 = cell8._address;
      data = { ...data , [add8] : "####"}
      let cell13 = rows[key1]._cells[13];
      let add13 = cell13._address;
      data = { ...data,[add13] : "####"}
      let cell27 = rows[key1]._cells[27];
      let add27 = getCellAddress(cell27, "dummy_address_27");
      data = { ...data,[add27] : "  "}
    }
        
  }
  if (
    rows[key1]._cells[0] != undefined &&
    rows[key1]._cells[0]._value.model.value == "$$vn"
  ) {
    vn_i = vn_i + 1;
    if (vn_i == 1) {
    let cell5 = rows[key1]._cells[5];
    let add5 = cell5._address;
    data = { ...data,[add5] : 'Vc +Vs +Vp'}
    let cell11 = rows[key1]._cells[11];
    let add11 = cell11._address;
    data = { ...data, [add11]: parseFloat(Vn.toFixed(2)) };
  }
  if (vn_i == 2){
    let cell5 = rows[key1]._cells[5];
    let add5 = cell5._address;
    data = { ...data,[add5] : 'Vc +Vs +Vp'}
    let cell11 = rows[key1]._cells[11];
    let add11 = cell11._address;
    data = { ...data, [add11]: parseFloat(Vn_min.toFixed(2)) };
  }
  
  }
    // Removed commented out code for cleanliness
    
    if (
      getSafeCell(rows[key1], 0) &&
      getSafeCell(rows[key1], 0)._value.model.value === "$$vr"
  ) {
    vu = rows[key1]._cells[17].value;
   
      vr_i = vr_i + 1;
      if (vr_i === 1) {
        vr_old = rows[key1]._cells[8].value;
          let cell8 = getSafeCell(rows[key1], 8);
          let cell17 = getSafeCell(rows[key1], 17);
          let cell29 = getSafeCell(rows[key1], 29);
          let cell16 = getSafeCell(rows[key1], 16);
          let add8 = cell8 ? cell8._address : null;
          let add29 = cell29 ? cell29._address : null;
          let add16 = cell16 ? cell16._address : null;
          let cell17_value = cell17 ? cell17.value : null;
  
          if (add8 && add29 && add16 && cell17_value !== null) {
              let vr = pi * Vn;
              if (vr < cell17_value) {
                  data = { ...data, [add16]: "<" };
                  data = { ...data, [add29]: "NG" };
              } else {
                  data = { ...data, [add16]: "≥" };
                  data = { ...data, [add29]: "OK" };
              }
              data = { ...data, [add8]: parseFloat(vr.toFixed(2)) };
              vr_new = parseFloat(vr.toFixed(2));
          } else {
              console.error("Error: Unable to retrieve necessary cell addresses or values for vr_i == 1");
          }
      }
  
      if (vr_i === 2) {
        vr_old_n = rows[key1]._cells[8].value;
          let cell8 = getSafeCell(rows[key1], 8);
          let cell17 = getSafeCell(rows[key1], 17);
          let cell29 = getSafeCell(rows[key1], 29);
          let cell16 = getSafeCell(rows[key1], 16);
          let add8 = cell8 ? cell8._address : null;
          let add29 = cell29 ? cell29._address : null;
          let add16 = cell16 ? cell16._address : null;
          let cell17_value = cell17 ? cell17.value : null;
  
          if (add8 && add29 && add16 && cell17_value !== null) {
              let vr_min = pi_min * Vn_min;
              if (vr_min < cell17_value) {
                  data = { ...data, [add16]: "<" };
                  data = { ...data, [add29]: "NG" };
              } else {
                  data = { ...data, [add16]: "≥" };
                  data = { ...data, [add29]: "OK" };
              }
              data = { ...data, [add8]: parseFloat(vr_min.toFixed(2)) };
              vr_new_n = parseFloat(vr_min.toFixed(2));
          } else {
              console.error("Error: Unable to retrieve necessary cell addresses or values for vr_i == 2");
          }
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
      }
    //   deleteRowsBetweenMarkers(worksheet);
      for (let key1 in rows) {
        if (
          rows[key1]._cells[0] != undefined &&
          rows[key1]._cells[0]._value.model.value == "$$b_str"  
        ) {
           bs_i = bs_i + 1;
          // Store the starting index for deletion
          if ( bs_i == 1 || bs_i == 2) {
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
  
        }
    }
  }
  
      for (let i = 0; i < (rows.length + 15); i++) {
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
    function getSafeCell(row, index) {
        try {
          if (row && row._cells && row._cells[index] && row._cells[index]._value && row._cells[index]._value.model) {
            return row._cells[index];
          }
        } catch (error) {
          console.error(`Error accessing cell at index ${index}:`, error);
        }
        return null;
      }
      function getCellAddress(cell, defaultAddress) {
        if (cell && cell._address) {
            return cell._address;
        } else {
            console.warn(`Cell address is null or undefined, using default: ${defaultAddress}`);
            return defaultAddress;
        }
    }
      function deleteRowsBetweenMarkers(worksheet) {
        for (let i = 1; i <= worksheet.rowCount; i++) {
          let row = worksheet.getRow(i);
          let cell = getSafeCell(row, 0);
      
          if (cell && cell._value.model.value === "$$b_str") {
            // Make the row with "$$b_str" and the next 3 rows blank
            for (let j = 0; j < 4; j++) {
              let currentRow = worksheet.getRow(i + j);
              currentRow.eachCell({ includeEmpty: true }, (cell) => {
                cell.value = null;
              });
            }
          }
        }
      }
      
      
  
    // function updatedata2(wkey, worksheet2,beamStresses) {
    //   if (!workbookData) return;
    //   if (!worksheet2) {
    //     throw new Error("No worksheets found in the uploaded file");
    //   }
    //     // Get the number of rows in the worksheet
    //     const rowCount = worksheet2.rowCount;
  
    //     // Access the last row
    //     const lastRowNumber = rowCount; // Row numbers are 1-based
    //     const lastRow = worksheet2.getRow(lastRowNumber);
        
  
    //     // Log the last row for debugging
    //     console.log(`Last row (${lastRowNumber}):`, lastRow);
  
    //     const thirdRow = worksheet2.getRow(4);
    //     const thirdRowCellValue = thirdRow.getCell(3).value;
    //     console.log(`Value of the cell in row 3, column 3: ${thirdRowCellValue}`); 
  
    //     const nextRowNumber = lastRowNumber + 1;
  
    //     // Access the next row
    //     const nextRow = worksheet2.getRow(nextRowNumber);
    
    //     // Populate the first cell with selectedName
    //     nextRow.getCell(1).value = beamStresses.BeamStress.DATA[0][1];
    //     nextRow.getCell(2).value = beamStresses.BeamStress.DATA[0][3];
    //     if (thirdRowCellValue == '-') {
    //     nextRow.getCell(3).value = '-';
    //     }
    //     else {
    //       nextRow.getCell(3).value = 'Girder';
    //     }
    //     nextRow.getCell(4).value = 'Tension';
    //     nextRow.getCell(5).value = selectedName;
    //     nextRow.getCell(8).value = changeSignAndFormat(beamStresses.BeamStress.DATA[0][12]);
    //     nextRow.getCell(9).value = changeSignAndFormat(beamStresses.BeamStress.DATA[0][15]);
    //     nextRow.getCell(10).value = changeSignAndFormat(beamStresses.BeamStress.DATA[0][13]);
    //     nextRow.getCell(11).value = changeSignAndFormat(beamStresses.BeamStress.DATA[0][14]); 
    //     nextRow.getCell(6).value = calculateAverage(nextRow.getCell(8).value, nextRow.getCell(10).value);
    //     nextRow.getCell(7).value = calculateAverage(nextRow.getCell(9).value, nextRow.getCell(11).value);
    //     nextRow.getCell(12).value = findMinValue([ nextRow.getCell(8),nextRow.getCell(9),nextRow.getCell(10),nextRow.getCell(11)]);
    //     nextRow.getCell(13).value = '0';
    //     if (nextRow.getCell(12).value > 0) {
    //         nextRow.getCell(14).value = 'OK';
    //     }
    //     else {
    //         nextRow.getCell(14).value = 'NG';
    //     }
    //     // Populate the second cell with Section part from beamStresses.data
    //     // Assuming beamStresses.data is an array and we want the first element
    
    //     // Log the next row for debugging
    //     console.log(`Next row (${nextRowNumber}):`, nextRow);
  
    //     // Perform operations on the last row
    //     // For example, you can get cell values, update them, etc.
    //     // const lastCellValue = lastRow.getCell(1).value;
    //     // lastRow.getCell(1).value = 'New Value';
  
    //     // Save changes to the last row (if necessary)
    //     if (thirdRowCellValue!='-') {
    //     const nextRowNumber2 = lastRowNumber + 2;
  
    //     // Access the next row
    //     const nextRow2 = worksheet2.getRow(nextRowNumber2);
    
    //     // Populate the first cell with selectedName
    //     nextRow2.getCell(1).value = beamStresses.BeamStress.DATA[0][1];
    //     nextRow2.getCell(2).value = beamStresses.BeamStress.DATA[0][5];
    //     nextRow2.getCell(3).value = 'Slab';
    //     nextRow2.getCell(4).value = 'Tension';
    //     nextRow2.getCell(5).value = selectedName;
    //     nextRow2.getCell(8).value = changeSignAndFormat(beamStresses.BeamStress.DATA[0][12]);
    //     nextRow2.getCell(9).value = changeSignAndFormat(beamStresses.BeamStress.DATA[0][15]);
    //     nextRow2.getCell(10).value = changeSignAndFormat(beamStresses.BeamStress.DATA[0][13]);
    //     nextRow2.getCell(11).value = changeSignAndFormat(beamStresses.BeamStress.DATA[0][14]); 
    //     nextRow2.getCell(6).value = calculateAverage(nextRow2.getCell(8).value, nextRow2.getCell(10).value);
    //     nextRow2.getCell(7).value = calculateAverage(nextRow2.getCell(9).value, nextRow2.getCell(11).value);
    //     nextRow2.getCell(12).value = findMinValue([ nextRow2.getCell(8),nextRow2.getCell(9),nextRow2.getCell(10),nextRow2.getCell(11)]);
    //     nextRow2.getCell(13).value = '0';
    //     if (nextRow2.getCell(12).value > 0) {
    //         nextRow2.getCell(14).value = 'OK';
    //     }
    //     else {
    //         nextRow2.getCell(14).value = 'NG';
    //     }
    //     // Populate the second cell with Section part from beamStresses.data
    //     // Assuming beamStresses.data is an array and we want the first element
    
    //     // Log the next row for debugging
    //     console.log(`Next row (${nextRowNumber2}):`, nextRow2);
    //   }
    //     lastRow.commit();
    // }
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
        let cell3Value = worksheet2.getRow(4).getCell(3).value;
    
        // Populate the first cell with selectedName
        nextRow.getCell(1).value = beamStresses.BeamStress.DATA[0][1];
        if (cell3Value == '-') {
          let value = beamStresses.BeamStress.DATA[0][3];

           // Remove the square brackets using a regular expression
           value = value.replace(/\[.*?\]/g, '');

           nextRow.getCell(2).value = value.trim();
        }
        else {
          nextRow.getCell(2).value = beamStresses.BeamStress.DATA[0][5];
        }
        if (cell3Value == '-') {
          nextRow.getCell(3).value = '-';
        }
        else {
          nextRow.getCell(3).value = 'Girder';
        }
         
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
        console.log(`Next row (${nextRowNumber}):`, nextRow);
        if(cell3Value !== '-') {
        const nextRowNumber2 = lastRowNumber + 2;
        const nextRow2 = worksheet2.getRow(nextRowNumber2);
    
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
        console.log(`Next row (${nextRowNumber2}):`, nextRow2);
      }
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
    
    async function updatedata3(wkey, worksheet3, name) {
      if (!workbookData) return;
      if (!worksheet3) {
        throw new Error("No worksheets found in the uploaded file");
      }
      console.log(mu_pos);
    
      const formatCell = (cell, value, bgColor = 'FFADD8E6', textColor = 'FF000000', bold = true) => {
        cell.value = value;
        cell.font = { bold: bold, color: { argb: textColor } };
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: bgColor } };
        cell.alignment = { vertical: 'middle', horizontal: 'center' };
        cell.border = {
          top: { style: 'thin' },
          left: { style: 'thin' },
          bottom: { style: 'thin' },
          right: { style: 'thin' },
        };
      };
    
      const formatNumberCell = (cell, value) => {
        cell.value = parseFloat(value).toFixed(3);
        cell.alignment = { vertical: 'middle', horizontal: 'center' };
        // cell.border = {
        //   top: { style: 'thin' },
        //   left: { style: 'thin' },
        //   bottom: { style: 'thin' },
        //   right: { style: 'thin' },
        // };
      };

      // const formatlistCell = (cell, value) => {
      //   cell.dataValidation={
      //     type: 'list',
      //     allowBlank: true,
      //     formulae: ['"One,Two,Three,Four"']
      //   }
      //   cell.DropList = ["asd","ada"];
      //   cell.alignment = { vertical: 'middle', horizontal: 'center' };
      //   // cell.border = {
      //   //   top: { style: 'thin' },
      //   //   left: { style: 'thin' },
      //   //   bottom: { style: 'thin' },
      //   //   right: { style: 'thin' },
      //   // };
      // };
    
      const formatResultCell = (cell, value) => {
        cell.value = value;
        cell.font = { color: { argb: value === "OK" ? 'FF00008B' : 'FFFF0000' } }; // Dark blue for OK, Red for NG
        cell.alignment = { vertical: 'middle', horizontal: 'center' };
        // cell.border = {
        //   top: { style: 'thin' },
        //   left: { style: 'thin' },
        //   bottom: { style: 'thin' },
        //   right: { style: 'thin' },
        // };
      };
      worksheet3.getColumn('A').width = 18;
      // Merge and format cells in the first two rows
      worksheet3.mergeCells('A1:A2');
      formatCell(worksheet3.getCell('A1'), 'Element');
      worksheet3.mergeCells('B1:G1');
      formatCell(worksheet3.getCell('B1'), name);
      worksheet3.mergeCells('B2:D2');
      formatCell(worksheet3.getCell('B2'), 'AASTHO');
      worksheet3.mergeCells('E2:G2');
      formatCell(worksheet3.getCell('E2'), 'Caltrans');
    
      // Fill specific cells
      formatNumberCell(worksheet3.getCell('B5'), mu_pos);
      formatNumberCell(worksheet3.getCell('E5'), mu_pos);
      formatNumberCell(worksheet3.getCell('B6'), mr_old_pos);
      formatNumberCell(worksheet3.getCell('E6'), mr_new_pos);
      formatResultCell(worksheet3.getCell('D6'), mr_old_pos < mu_pos ? "NG" : "OK");
      formatResultCell(worksheet3.getCell('G6'), mr_new_pos < mu_pos ? "NG" : "OK");
    
      // Array of rows to merge and format
      const rowsToMerge = [3, 7, 11, 14, 17, 21, 28];
      const contentForRows = [
        "1. Factored Resistance: Positive ",
        "2. Factored Resistance: Negative ",
        "3. Maximum spacing for transverse reinforcement: Maximum shear case ",
        "4. Maximum spacing for transverse reinforcement: Minimum shear case ",
        "5. Crack Check ",
        "6. Shear Resistance parameters : Maximum ",
        "7. Shear Resistance parameters : Minimum "
      ];
    
      // Merge and format specified rows
      rowsToMerge.forEach((row, index) => {
        worksheet3.mergeCells(`A${row}:G${row}`);
        formatCell(worksheet3.getCell(`A${row}`), contentForRows[index], 'FFD3D3D3');
      });
    
      // Rows to skip while merging B-C and E-F
      const rowsToSkip = [1, 2, 3, 7, 11, 14, 17, 21, 28];
    
      // Merge cells B and C, E and F up to row 29, skipping specified rows
      for (let i = 1; i <= 34; i++) {
        if (!rowsToSkip.includes(i)) {
          worksheet3.mergeCells(`B${i}:C${i}`);
          worksheet3.mergeCells(`E${i}:F${i}`);
        }
      } 
      // Additional cells to format
      formatCell(worksheet3.getCell('A4'), "ϕ");
      formatCell(worksheet3.getCell('A5'), "Mu(kips·in)");
      formatCell(worksheet3.getCell('A6'), 'Mr(kips·in)');
      formatCell(worksheet3.getCell('A8'), 'ϕ');
      formatCell(worksheet3.getCell('A9'), 'Mu(kips·in)');
      formatCell(worksheet3.getCell('A10'), 'Mr(kips·in)');
      formatCell(worksheet3.getCell('A12'), 'Smax(in)');
      formatCell(worksheet3.getCell('A13'), 'S(in)');
      formatCell(worksheet3.getCell('A15'), 'Smax(in)');
      formatCell(worksheet3.getCell('A16'), 'S(in)');
      formatCell(worksheet3.getCell('A18'), 'dc(in)');
      formatCell(worksheet3.getCell('A19'), 'smax(in)');
      formatCell(worksheet3.getCell('A20'), 's(in)');
      formatCell(worksheet3.getCell('A22'), 'β');
      formatCell(worksheet3.getCell('A23'), 'θ');
      formatCell(worksheet3.getCell('A24'), 'Av(in²)');
      formatCell(worksheet3.getCell('A25'), 'Av,req(in²)');
      formatCell(worksheet3.getCell('A26'), 'Vu(kips)');
      formatCell(worksheet3.getCell('A27'), 'Vr(kips)');
      formatCell(worksheet3.getCell('A22'), 'β');
      formatCell(worksheet3.getCell('A23'), 'θ');
      formatCell(worksheet3.getCell('A24'), 'Av(in²)');
      formatCell(worksheet3.getCell('A25'), 'Av,req(in²)');
      formatCell(worksheet3.getCell('A26'), 'Vu(kips)');
      formatCell(worksheet3.getCell('A27'), 'Vr(kips)');
      formatCell(worksheet3.getCell('A29'), 'β');
      formatCell(worksheet3.getCell('A30'), 'θ');
      formatCell(worksheet3.getCell('A31'), 'Av(in²)');
      formatCell(worksheet3.getCell('A32'), 'Av,req(in²)');
      formatCell(worksheet3.getCell('A33'), 'Vu(kips)');
      formatCell(worksheet3.getCell('A34'), 'Vr(kips)');
      formatNumberCell(worksheet3.getCell('B4'), 1);
      // formatlistCell(worksheet3.getCell('B4'),1);
      formatNumberCell(worksheet3.getCell('E4'), phi_new_m);
      formatNumberCell(worksheet3.getCell('B8'), 1);
      formatNumberCell(worksheet3.getCell('E8'), phi_new_n);
      formatNumberCell(worksheet3.getCell('B9'), mu_neg);
      formatNumberCell(worksheet3.getCell('B10'), mr_old_neg);
      formatNumberCell(worksheet3.getCell('E9'), mu_neg);
      formatNumberCell(worksheet3.getCell('E10'), mr_new_neg);
      formatResultCell(worksheet3.getCell('D10'), mr_old_neg < mu_neg ? "NG" : "OK");
      formatResultCell(worksheet3.getCell('G10'), mr_new_neg < mu_neg ? "NG" : "OK");
      formatNumberCell(worksheet3.getCell('B12'), sm_old);
      formatNumberCell(worksheet3.getCell('E12'), sm_new);
      formatNumberCell(worksheet3.getCell('B13'), s_m);
      formatNumberCell(worksheet3.getCell('E13'), s_m);
      formatResultCell(worksheet3.getCell('D13'), sm_old < s_m ? "NG" : "OK");
      formatResultCell(worksheet3.getCell('G13'), sm_new < s_m ? "NG" : "OK");
      formatNumberCell(worksheet3.getCell('B15'), sn_old);
      formatNumberCell(worksheet3.getCell('E15'), sn_new);
      formatNumberCell(worksheet3.getCell('B16'), s_n);
      formatNumberCell(worksheet3.getCell('E16'), s_n);
      formatResultCell(worksheet3.getCell('D16'), sn_old < s_n ? "NG" : "OK");
      formatResultCell(worksheet3.getCell('G16'), sn_new < s_n ? "NG" : "OK");
      formatNumberCell(worksheet3.getCell('B18'), dc_old);
      formatNumberCell(worksheet3.getCell('B19'), smax_old);
      formatNumberCell(worksheet3.getCell('B20'), suse);
      formatResultCell(worksheet3.getCell('D20'), smax_old < suse ? "NG" : "OK");
      formatResultCell(worksheet3.getCell('G20'), smax_new < suse ? "NG" : "OK");
      formatNumberCell(worksheet3.getCell('E18'), dc_new);
      formatNumberCell(worksheet3.getCell('E19'), smax_new);
      formatNumberCell(worksheet3.getCell('E20'), suse);
      formatNumberCell(worksheet3.getCell('B22'), beta_mo);
      formatNumberCell(worksheet3.getCell('B23'), theta_mo);
      formatNumberCell(worksheet3.getCell('E22'), beta_m);
      formatNumberCell(worksheet3.getCell('E23'), theta_m);
      formatNumberCell(worksheet3.getCell('B24'), Av_f);
      formatNumberCell(worksheet3.getCell('B25'), Avr_old);
      formatNumberCell(worksheet3.getCell('E24'), Av_f);
      formatNumberCell(worksheet3.getCell('E25'), Avr_new);
      formatResultCell(worksheet3.getCell('D25'), Avr_old > Av_f ? "NG" : "OK");
      formatResultCell(worksheet3.getCell('G25'), Avr_new > Av_f ? "NG" : "OK");
      formatNumberCell(worksheet3.getCell('B26'), vu);
      formatNumberCell(worksheet3.getCell('B27'), vr_old);
      formatNumberCell(worksheet3.getCell('E26'), vu);
      formatNumberCell(worksheet3.getCell('E27'), vr_new);
      formatResultCell(worksheet3.getCell('D27'), vu > vr_old ? "NG" : "OK");
      formatResultCell(worksheet3.getCell('G27'), vu > vr_new ? "NG" : "OK");
      formatNumberCell(worksheet3.getCell('B29'), beta_no);
      formatNumberCell(worksheet3.getCell('B30'), theta_no);
      formatNumberCell(worksheet3.getCell('E29'), beta_n);
      formatNumberCell(worksheet3.getCell('E30'), theta_n);
      formatNumberCell(worksheet3.getCell('B31'), Av_f);
      formatNumberCell(worksheet3.getCell('B32'), Avr_old);
      formatNumberCell(worksheet3.getCell('E31'), Av_f);
      formatNumberCell(worksheet3.getCell('E32'), Avr_new);
      formatResultCell(worksheet3.getCell('D32'), Avr_old > Av_f ? "NG" : "OK");
      formatResultCell(worksheet3.getCell('G32'), Avr_new > Av_f ? "NG" : "OK");
      formatNumberCell(worksheet3.getCell('B33'), vu);
      formatNumberCell(worksheet3.getCell('B34'), vr_old);
      formatNumberCell(worksheet3.getCell('E33'), vu);
      formatNumberCell(worksheet3.getCell('E34'), vr_new);
      formatResultCell(worksheet3.getCell('D34'), vu > vr_old ? "NG" : "OK");
      formatResultCell(worksheet3.getCell('G34'), vu > vr_new ? "NG" : "OK");
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
      }
      showLc(allData);
      return null;
    }
    function showLc(lc) {
      console.log(lc);
      item.delete("1");
      let newKey = 1;
      for (let key in lc) {
        if (lc[key].ACTIVE === "SERVICE") {
          item.set(lc[key].NAME, newKey.toString());
          newKey++;
        }
      }
  
      setItem(item);
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
        setButtonText('Creating...');
      // fetchLc();
      const combArray = Object.values(lc);
      let beamStresses;
      let beamStresses_box;
      if (combArray.length === 0) {
        enqueueSnackbar("Please Define Load Combination", {
          variant: "error",
          anchorOrigin: {
            vertical: "top",
            horizontal: "center",
          },
          action,
        });
        return;
      }
      console.log(SelectWorksheets);
      console.log(SelectWorksheets2);
      let numberPart = parseInt(matchedParts[0].numberPart, 10);
      let letterPart = matchedParts[0].letterPart;
      let name = numberPart + "_" + letterPart;
      console.log(selectedName)
      console.log(name);
      const concatenatedValue_cbc = `${selectedName}(CBC)`;
      const concatenatedValue_cbc_max = `${selectedName}(CBC:max)`;
      const concatenatedValue_cb = `${selectedName}(CB)`;
      const concatenatedValue_cd_max = `${selectedName}(CB:max)`;
      const concatenatedValue_cbr = `${selectedName}(CBR)`;
      const concatenatedValue_cbr_max = `${selectedName}(CBR:max)`;
      const concatenatedValue_cbsc = `${selectedName}(CBSC)`;
      const concatenatedValue_cbsc_max = `${selectedName}(CBSC:max)`;
      let box_stresses = {
        "Argument": {
            "TABLE_NAME": "BeamStress",
            "TABLE_TYPE": "BEAMSTRESS",
            "EXPORT_PATH": "C:\\MIDAS\\Result\\Output.JSON",
            "STYLES": {
                "FORMAT": "Fixed",
                "PLACE": 12
            },
            "COMPONENTS": [
                "Elem",
                "Load",
                "Part",
                "Axial",
                "Shear-y",
                "Shear-z",
                "Bend(+y)",
                "Bend(-y)",
                "Bend(+z)",
                "Bend(-z)",
                "Cb(min/max)",
                "Cb1(-y+z)",
                "Cb2(+y+z)",
                "Cb3(+y-z)",
                "Cb4(-y-z)"
            ],
            "NODE_ELEMS": {
                "KEYS": [
                  numberPart
                ]
            },
            "LOAD_CASE_NAMES": [
              selectedName,
              concatenatedValue_cbc,
              concatenatedValue_cbc_max,
              concatenatedValue_cb,
              concatenatedValue_cd_max,
              concatenatedValue_cbr,
              concatenatedValue_cbr_max,
              concatenatedValue_cbsc,
              concatenatedValue_cbsc_max
            ],
            "PARTS": [
               `Part ${letterPart}`
            ]
        }
    };
    console.log(box_stresses);
  
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
                concatenatedValue_cbc,
                concatenatedValue_cbc_max,
                concatenatedValue_cb,
                concatenatedValue_cd_max,
                concatenatedValue_cbr,
                concatenatedValue_cbr_max,
                concatenatedValue_cbsc,
                concatenatedValue_cbsc_max
            ],
            "PARTS": [
                  `Part ${letterPart}`
            ]
        }
    };
    console.log(stresses);
    try {
      beamStresses = await midasAPI("POST", "/post/table", stresses);
      
      if (beamStresses.message === '') {
        beamStresses_box = await midasAPI("POST", "/post/table", box_stresses);
      }
    
      // setBeamStresses(beamStresses);
      console.log(beamStresses);
      console.log(beamStresses_box);
    } catch (error) {
      console.error("Error fetching beam stresses:", error);
    }
    let type;
      for (let wkey in SelectWorksheets) {
        updatedata(wkey, SelectWorksheets[wkey]);
      }    
      for (let wkey in SelectWorksheets2) {
        // Check if beamStresses is not null or undefined and has keys
        if (beamStresses.BeamStress && beamStresses.BeamStress.DATA) {
          updatedata2(wkey, SelectWorksheets2[wkey], beamStresses);
        } else {
          updatedata2(wkey, SelectWorksheets2[wkey], beamStresses_box);
        }
      }
      console.log(beamStresses);
      console.log(beamStresses_box);
      
      for (let wkey in SelectWorksheets3) {
          updatedata3(wkey, SelectWorksheets3[wkey],name);
        }
      if (!workbookData) return;
      try {
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
    
        // Show success notification
        enqueueSnackbar("Output file generated successfully!", {
          variant: "success",
          anchorOrigin: {
            vertical: "top",
            horizontal: "center",
          },
          action,
        });
        if (fileInputRef.current) {
            fileInputRef.current.value = null;
          }
          setButtonText('Create Report');
        //   await handleFileDownload;
      } catch (error) {
        // Show error notification
        enqueueSnackbar(`Error generating output file: ${error.message}`, {
          variant: "error",
          anchorOrigin: {
            vertical: "top",
            horizontal: "center",
          },
          action,
        });
        setButtonText('Create Report');
      }
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
      <Panel width={510} height={470} marginTop={3} padding={2} variant="shadow2">
        <Panel width={480} height={200} marginTop={0} variant="shadow2">
        <div style={{ marginTop: "8px" }}>
          <Grid container>
            <Grid item xs={9}>
              <Typography variant="h1">
                {" "}
                Casting Method
              </Typography>
            </Grid>
            <Grid item xs={3}>
              <Typography variant="h1"> (5.5.4.2)</Typography>
            </Grid>
          </Grid>
          {/* <RadioGroup
            margin={1}
            onChange={(e) => setSp(e.target.value)} // Update state variable based on user selection
            value={cast} // Bind the state variable to the RadioGroup
            text=""
          > */}
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
                width: "235px",
              }}
            >
              <Radio name="CA (2.5 inches)" value="ca2" checked={cvr === "ca2"} />
              <Radio
                name="AASHTO LFRD"
                value="aa2"
                checked={cvr === "aa2"}
              />
            </div>
          </RadioGroup>
        </div>
        </Panel>
        <Panel width={480} height={200} marginTop={1} variant="shadow2">
        <div style={{ marginTop: "6px" }}>
          <Grid container>
            <Grid item xs={3}>
              <Typography variant="h1">
                {" "}
                Load Combination for SLS (Permanent Loads)
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
            marginTop: "8px",
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
            <Grid item xs={4}>
            {/* <img
          src={Image} // Use the imported image
          alt="Description of the image"
          style={{ width: '100%', height: 'auto' }} // Inline styles for responsiveness
        /> */}
            </Grid>
          </Grid>
          {/*  */}
          <Grid container direction="row">
            <Grid item xs={9}>
              <Typography>Maximum aggregate size(ag) (in inches)</Typography>
            </Grid>
            <Grid item xs={3}>
              <TextField
                value={ag}
                onChange={handleAgChange}
                placeholder=""
                //   title="Maximum aggregate size(ag)"
                width="100px"
              />
            </Grid>
            </Grid>
            <Grid container direction="row">
            <Grid item xs={9} marginTop={0.5}>
              <Typography size="small">
              <span dangerouslySetInnerHTML={{ __html: 'Maximum distance between the layers of longitudinal crack control reinforcement (s<sub>xe</sub>) (in inches)' }} />
              </Typography>
            </Grid>
            <Grid item xs={3} marginTop={0.5}>
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
        </Panel>
        <div
          style={{
            display: "flex",
            justifyContent: "space-between",
            margin: "0px",
            marginTop: "10px",
            marginBottom: "30px",
          }}
        >
          {/* {Buttons.NormalButton("contained", "Import Report", () => importReport())} */}
          {/* {Buttons.MainButton("contained", "Update Report", () => updatedata())}  */}
          {Buttons.MainButton("contained", buttonText, handleFileDownload)}
          {check && <AlertDialogModal />}
        </div>
      </Panel>
    );
  };