import { DropList, Grid, Panel, Typography, VerifyUtil } from '@midasit-dev/moaui';
import { Radio, RadioGroup } from "@midasit-dev/moaui";
import React, { useState } from 'react';
import * as Buttons from "./Components/Buttons";
import ExcelJS from 'exceljs';
import AlertDialogModal from './AlertDialogModal';
import { midasAPI } from "./Function/Common";
import { enqueueSnackbar } from 'notistack';
import { ThetaBeta1 } from './Function/ThetaBeta';
import { TextField } from '@midasit-dev/moaui'

export const Updatereport = () => {
    const [workbookData, setWorkbookData] = useState(null);
    const [sheetData, setSheetData] = useState([]);
    const [sheetName, setSheetName] = useState('');
    const [cast, setCast] = useState("inplace");
    const [sp, setSp] = useState("ca1");
    const [cvr, setCvr] = useState("ca2");
    const [value, setValue] = useState(1);
    const [SelectWorksheets, setWorksheet] = useState({})
    const [Lc, setLc] = useState({});
    const [item, setItem] = useState(new Map([['Select Load Combination', 1]]))
    const [check, setCheck] = useState(false);
    function onChangeHandler(event) {
        setValue(event.target.value);
    }
    let items = new Map([]);

    const handleFileUpload = (event) => {
        const file = event.target.files[0];
        const reader = new FileReader();
        reader.onload = async (e) => {
            //    await Promise.all([fetchLc()]) 
            await fetchLc();
            // console.log(lc)
            // showLc(lc)

            try {
                let buffer = e.target.result;
                let workbook = new ExcelJS.Workbook();
                await workbook.xlsx.load(buffer);

                // Get the first worksheet
                // console.log(workbook);
                let worksheet;
                console.log(workbook);
                for (let key in workbook.worksheets) {
                    const regex = /^[0-9]+_[A-Z]$/;
                    if (regex.test(workbook.worksheets[key].name)) {
                        // console.log(workbook.worksheets[key].name)
                        worksheet = workbook.worksheets[key];
                        setWorksheet(prevstate => ({
                            ...prevstate, [key]: workbook.worksheets[key]
                        }));
                    }

                }
                if (!worksheet) {
                    throw new Error('No worksheets found in the uploaded file');
                }
                else {
                    console.log(worksheet);
                    let cellvalue = worksheet._rows[2]._cells[2]._value.value;
                    if (cellvalue != 'AASHTO-LRFD2017') {
                        alert();
                    }
                }
                setWorkbookData(workbook);
                setSheetName(worksheet.name);

            } catch (error) {
                console.error('Error reading file:', error);
                alert('Error reading file. Please make sure the file is a valid Excel file.');
            }
        };

        reader.readAsArrayBuffer(file);
    };
    const [ag, setAg] = useState('');
    const [sg, setSg] = useState('');
   
  
    const handleAgChange = (event) => {
      setAg(event.target.value);
    };
  
    const handleSgChange = (event) => {
      setSg(event.target.value);
    };
    console.log(ag);
    console.log(sg);
    function updatedata(wkey, worksheet) {
        if (!workbookData) return;
        if (!worksheet) {
            throw new Error('No worksheets found in the uploaded file');
        }
        
     

        let rows = worksheet._rows;
        let mn;
        let phi;
        let mr;
        let dv;
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
        let Vu1; let Vu2;
        let beta1;
        let Vc;
        let Vc1;
        let dc;
        let storedValues = {};

        const updateCellValue = (rowKey, cellIndex, value) => {
            // Update the local rows structure
            let cellAddress = indexToLetter(cellIndex) + (parseInt(rowKey) + 1);
            if (!rows[rowKey]._cells[cellIndex]) {
                // Create a dummy cell if it's missing
                rows[rowKey]._cells[cellIndex] = {
                    _value: {
                        model: {
                            value: value,
                            address: cellAddress // Assign the value
                        }
                    },
                    _address: cellAddress
                };
            } else {
                // If the cell is present, update its value
                rows[rowKey]._cells[cellIndex]._value.model.value = value;
            }
        
            // Update the worksheet
            worksheet.getCell(cellAddress).value = value;
        };

        for (let key1 in rows) {           // to traverse all the rows of excel sheet 

            if (rows[key1]._cells[0] != undefined && rows[key1]._cells[0]._value.model.value == '$$Mn') {
                let location = rows[key1]._cells[19]._value.model.address;
                let value = rows[key1]._cells[19]._value.model.value
                data = { ...data, [location]: value };
                mn = value;
            }

            // to get Phi row
            if (rows[key1]._cells[0] != undefined && rows[key1]._cells[0]._value.model.value == '$$Phi') {
                let location = rows[key1]._cells[5]._value.model.address;
                if (cast === "inplace") {
                    data = { ...data, [location]: 0.95 };
                    phi = 0.95
                }
                else {
                    data = { ...data, [location]: 1 };
                    phi = 1;
                }
                console.log(phi);
            }
        

            if (rows[key1]._cells[0] != undefined && rows[key1]._cells[0]._value.model.value == '$$Mr') {
                let location = rows[key1]._cells[5]._value.model.address;

                let mu = rows[key1]._cells[17]._value.model.value;
                mr = Number(mn) * Number(phi);
                data = { ...data, [location]: mr };

                // location of oK
                if (mr < Number(mu)) {
                    let location1 = rows[key1]._cells[29]._value.model.address;
                    let location2 = rows[key1]._cells[13]._value.model.address;
                    data = { ...data, [location1]: 'NG' };
                    data = { ...data, [location2]: '<' };
                }
            }
            if (rows[key1]._cells[0] != undefined && rows[key1]._cells[0]._value.model.value == '$$dv') {
                dv = rows[key1]._cells[4]._value.model.value;
            }
            if (rows[key1]._cells[0] != undefined && rows[key1]._cells[0]._value.model.value == '$$sm') {  
                if (sp === "ca1") {
                    let add1 = rows[key1]._cells[6]._value.model.address;
                    data = { ...data, [add1]: 'Min[0.8dv, 18.0(in.)]' };
                    let add2 = rows[key1]._cells[13]._value.model.address;
                    // let val=rows[key1]._cells[13]._value.model.value;
                    if (0.8 * dv >= 18) {
                        data = { ...data, [add2]: 18 };
                    }
                    else {
                        data = { ...data, [add2]: 0.8 * dv };
                    }
                }
            }
            // if (rows[key1]._cells[0] != undefined && rows[key1]._cells[0]._value.model.value == '$$dc') {
            //     if (cvr === "ca2") {
            //         let add1 = rows[key1]._cells[8]._value.model.address;
            //         let add2 = rows[key1]._cells[11]._value.model.address;
            //         let val2 = rows[key1]._cells[11]._value.model.value + 3.6 - 5;
            //         let val3 = rows[key1]._cells[21]._value.model.value;
            //         let add4 = rows[key1]._cells[29]._value.model.address;
            //         let add5 = rows[key1]._cells[17]._value.model.address;
            //         if (val2 < 0) {
            //             val2 = 0.0;
            //         }
            //         if (val2 < val3) {
            //             data = { ...data, [add4]: 'NG' };
            //             data = { ...data, [add5]: '<' };
            //         }
            //         data = { ...data, [add1]: '2*2.5' };
            //         data = { ...data, [add2]: val2 };
            //     }
            // }
            if (rows[key1]._cells[0] != undefined && rows[key1]._cells[0]._value.model.value == '$$dc') {
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
                        if (rows[nextKey1]._cells[0]._value.model.value == '$$B') {
                            // Store the value and address of cell in column 13 for $$B row
                            column15Address = rows[nextKey1]._cells[15]._value.model.address;
                            column15Value = rows[nextKey1]._cells[15]._value.model.value;
                            column15Value_new = (1+((1*2.5)/(((1/((column15Value -1)/ 1.8))+1.26)-1.75))); 
                            column15Value_new = Math.round(column15Value_new);
                            console.log(column15Value_new);
                            storedValues[column15Address] = column15Value;
                            data = { ...data, [column15Address]: column15Value_new };
                        }
                        if (rows[nextKey1]._cells[0]._value.model.value == '$$d-c') {
                            // Store the value and address of cell in column 9 for $$dc row
                             column9Value = rows[nextKey1]._cells[9]._value.model.value;
                             column9Address = rows[nextKey1]._cells[9]._value.model.address;
                             column9Value_new = column9Value - column9Value + 2.5;
                             console.log(column9Value_new);
                             storedValues[column9Address] = column9Value;
                             data = { ...data, [column9Address]: column9Value_new };
                             break;
                        }
                    }
                    nextKey1++;
                }
                console.log(storedValues);
                val2_new = (((val2 + 3.6)*column15Value)/column15Value_new) - 5
            
                if (val2_new < 0) {
                    val2_new = 0.0;
                }
            
                if (val2_new < val3) {
                    data = { ...data, [add4]: 'NG' };
                    data = { ...data, [add5]: '<' };
                }
            
                data = { ...data, [add1]: '2*2.5' };
                data = { ...data, [add2]: val2_new };
             }
            }

            if (rows[key1]._cells[0] != undefined && rows[key1]._cells[0]._value.model.value == '$$Avm') {
                Avm = rows[key1]._cells[12]._value.model.value;
            }
            if (rows[key1]._cells[0] != undefined && rows[key1]._cells[0]._value.model.value == '$$Av') {
                Av = rows[key1]._cells[5]._value.model.value;
            }
            if (rows[key1]._cells[0] != undefined && rows[key1]._cells[0]._value.model.value == '$$Mmax') {
                Mmax = rows[key1]._cells[15]._value.model.value;
            }
            if (rows[key1]._cells[0] != undefined && rows[key1]._cells[0]._value.model.value == '$$Mmin') {
                Mmin = rows[key1]._cells[15]._value.model.value;
            }
            if (rows[key1]._cells[0] != undefined && rows[key1]._cells[0]._value.model.value == '$$Ag') {
                Ag = rows[key1]._cells[14]._value.model.value;
                St = rows[key1]._cells[24]._value.model.value;
            }
            if (rows[key1]._cells[0] != undefined && rows[key1]._cells[0]._value.model.value == '$$Sb') {
                Sb = rows[key1]._cells[24]._value.model.value;
            }
            if (rows[key1]._cells[0] != undefined && rows[key1]._cells[0]._value.model.value == '$$Nmax') {
                Nmax = rows[key1]._cells[15]._value.model.value;
            }
            if (rows[key1]._cells[0] != undefined && rows[key1]._cells[0]._value.model.value == '$$Nmin') {
                Nmin = rows[key1]._cells[15]._value.model.value;
            }
            if (rows[key1]._cells[0] != undefined && rows[key1]._cells[0]._value.model.value == '$$E') {
                E = rows[key1]._cells[9]._value.model.value;
                fc = rows[key1]._cells[2]._value.model.value;
            }
            if (rows[key1]._cells[0] != undefined && rows[key1]._cells[0]._value.model.value == '$$Vu1') {
                Vu1 = rows[key1]._cells[10]._value.model.value;
            }
            if (rows[key1]._cells[0] != undefined && rows[key1]._cells[0]._value.model.value == '$$Vu2') {
                Vu2 = rows[key1]._cells[10]._value.model.value;
            }
            if (rows[key1]._cells[0] != undefined && rows[key1]._cells[0]._value.model.value == '$$Vc') {
                Vc = rows[key1]._cells[9]._value.model.value;
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
        // // console.log(Avm, Av, worksheet);
        // if (Av >= Avm) {
        let Ecm = (-1 * Number(Mmax) / Number(St) + Number(Nmax) / Number(Ag)) / Number(E);
        let Ecn = (-1 * Number(Mmin) / Number(St) + Number(Nmin) / Number(Ag)) / Number(E);
        let Etm = (-1 * Number(Mmax) / Number(Sb) + Number(Nmax) / Number(Ag)) / Number(E);
        let Etn = (-1 * Number(Mmin) / Number(Sb) + Number(Nmin) / Number(Ag)) / Number(E);
        // console.log(Vu1,Vu2,fc)
        let a1 = Number(Vu1) / Number(fc);  
        let a2 = Number(Vu2) / Number(fc);
        let Exm = (Ecm + Etm) / 2; let Exn = (Ecn + Etn) / 2;
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
        let beta ;
        function indexToLetter(index) {
            // Convert a zero-based index to a letter (A, B, C, ..., Z, AA, AB, etc.)
            let letter = '';
            while (index >= 0) {
                letter = String.fromCharCode((index % 26) + 65) + letter;
                index = Math.floor(index / 26) - 1;
            }
            return letter;
        }
        for (let key1 in rows) {    
            if (rows[key1]._cells[0] != undefined && rows[key1]._cells[0]._value.model.value == '$$strm') {
                let add1 = rows[key1]._cells[2]._value.model.address;
                let add2 = rows[key1]._cells[8]._value.model.address;
                data = { ...data, [add1]: 'Calculation for β and θ' };
                data = { ...data, [add2]: '' };
            }  
            // if (rows[key1]._cells[0] !== undefined && rows[key1]._cells[0]._value !== undefined) {
            //     let cellValue = rows[key1]._cells[0]._value.model.value;
        
            //     if (cellValue === '$$strm1') {
            //         startBlanking = true;
            //     }
        
            //     if (startBlanking) {
            //         // Blank all cells in the row
            //         for (let cell of rows[key1]._cells) {
            //             if (cell._value !== undefined && cell._value.model !== undefined) {
            //                 cell._value.model.value = ''; // Blank the cell value
            //             }
            //         }
            //     }
        
            //     if (cellValue === '$$fpo') {
            //         break; // Stop blanking when the end marker is found
            //     }
            // }
            // let cell;
            // if (rows[key1]._cells[0] != undefined && rows[key1]._cells[0]._value.model.value == '$$strm1') {
            //     console.log(rows[key1])
            
            //     for (let i = 1; i <= 50; i++) {
            //      cell = rows[key1]._cells[i];
            //         if (cell && cell.model) {
            //             cell.model = {};
            //             cell.model.value = ' '; // Set model to an empty object
            //         }
            //     }
        
            //     // Run another while loop to make corresponding column cells empty until $$fpo is found
            //     let nextKey = parseInt(key1) + 1;
            //     while (rows[nextKey] !== undefined) {
            //         console.log(rows[nextKey]);
            //         if (rows[nextKey]._cells[0] != undefined && rows[nextKey]._cells[0]._value.model.value == '$$fpo') {
            //             break; // Stop the loop when $$fpo is found
            //         }       
            //         // Blank all cells in the current row of the while loop
            //         for (let i = 1; i <= 50; i++) {
            //             let cell = rows[nextKey]._cells[i];
            //             if (cell && cell.model) {
            //                 cell.model = { }; 
            //                 cell.model.value = ' ';
            //             }
            //         }
        
        
            //         nextKey++;
            //     }
          
            let cell;
            let add15value = dv < sg ? dv : sg;
            let sxe = ((add15value*1.38)/(ag + 0.63))
// if (rows[key1]._cells[0] != undefined && rows[key1]._cells[0]._value.model.value == '$$strm1') {
//     console.log(rows[key1]);
   
//     for (let i = 1; i <= 50; i++) {
//         // Check if the cell is present
//         if (!rows[key1]._cells[i]) {
//             // Create a dummy cell if it's missing
//             rows[key1]._cells[i] = {
//                 _value: {model: {
//                     value: 'dummy',
//                     address: indexToLetter(i) + (parseInt(key1) + 1)// Assign a dummy value
//                     // Add other necessary properties if required
                    
//                 }},
                
//                 _address : indexToLetter(i) + (parseInt(key1) + 1)
//             };
//         } else {
//             // If the cell is present, clear its value
//             cell = rows[key1]._cells[i];
//             if (cell && cell.model) {
//                 cell.model = {};
//                 cell.model.value = ' ';   // Clear the value
//             }
//         }
//     }

//     // Run another while loop to make corresponding column cells empty until $$fpo is found
//     let nextKey = parseInt(key1) + 1;
//     while (rows[nextKey] !== undefined) {
//         console.log(rows[nextKey]);
//         if (rows[nextKey]._cells[0] != undefined && rows[nextKey]._cells[0]._value.model.value == '$$fpo') {
//             break; // Stop the loop when $$fpo is found
//         }
//         // Blank all cells in the current row of the while loop
//         for (let i = 1; i <= 50; i++) {
//             // Check if the cell is present
//             if (!rows[nextKey]._cells[i]) {
//                 // Create a dummy cell if it's missing
//                 rows[nextKey]._cells[i] = {
//                     _value: {
//                         model: {
//                             value: 'dummy', // Assign a dummy value
//                             address: indexToLetter(i) + (nextKey + 1)// Add other necessary properties if required
//                         }
//                     },                   
//                     _address : indexToLetter(i) + (nextKey + 1)
//                 };
//             } else {
//                 // If the cell is present, clear its value
//                 cell = rows[nextKey]._cells[i];
//                 if (cell && cell.model) {
//                     cell.model = {};
//                     cell.model.value = ' '; // Clear the value
//                 }
//             }
//         }

//         nextKey++;
//     }
// }
// if (rows[key1]._cells[0] != undefined && rows[key1]._cells[0]._value.model.value == '$$strm1') {
//     console.log(rows[key1]);
   
//     for (let i = 1; i <= 50; i++) {
//         // Check if the cell is present
//         if (!rows[key1]._cells[i]) {
//             // Create a dummy cell if it's missing
//             rows[key1]._cells[i] = {
//                 _value: {model: {
//                     value: 'dummy',
//                     address: indexToLetter(i) + (parseInt(key1) + 1)// Assign a dummy value
//                     // Add other necessary properties if required
                    
//                 }},
                
//                 _address : indexToLetter(i) + (parseInt(key1) + 1)
//             };
//         } else {
//             // If the cell is present, clear its value
//             cell = rows[key1]._cells[i];
//             if (cell && cell.model) {
//                 cell.model = {};
//                 cell.model.value = ' ';   // Clear the value
//             }
//         }
//     }

//     // Run another while loop to make corresponding column cells empty until $$fpo is found
//     let nextKey = parseInt(key1) + 1;
//     while (rows[nextKey] !== undefined) {
//         console.log(rows[nextKey]);
//         if (rows[nextKey]._cells[0] != undefined && rows[nextKey]._cells[0]._value.model.value == '$$fpo') {
//             break; // Stop the loop when $$fpo is found
//         }
//         // Blank all cells in the current row of the while loop
//         for (let i = 1; i <= 50; i++) {
//             // Check if the cell is present
//             if (!rows[nextKey]._cells[i]) {
//                 // Create a dummy cell if it's missing
//                 rows[nextKey]._cells[i] = {
//                     _value: {
//                         model: {
//                             value: 'dummy', // Assign a dummy value
//                             address: indexToLetter(i) + (nextKey + 1)// Add other necessary properties if required
//                         }
//                     },                   
//                     _address : indexToLetter(i) + (nextKey + 1)
//                 };
//             } else {
//                 // If the cell is present, clear its value
//                 cell = rows[nextKey]._cells[i];
//                 if (cell && cell.model) {
//                     cell.model = {};
//                     cell.model.value = ' '; // Clear the value
//                 }
//             }
//         }

//         nextKey++;
//     }
// }


if (rows[key1]._cells[0] != undefined && rows[key1]._cells[0]._value.model.value == '$$strm1') {
    console.log(rows[key1]);

    for (let i = 1; i <= 50; i++) {
        // Check if the cell is present and update its value
        if (!rows[key1]._cells[i]) {
            updateCellValue(key1, i, 'dummy');
        } else {
            updateCellValue(key1, i, ' ');
        }
    }

    // Run another while loop to make corresponding column cells empty until $$fpo is found
    let nextKey = parseInt(key1) + 1;
    while (rows[nextKey] !== undefined) {
        console.log(rows[nextKey]);
        if (rows[nextKey]._cells[0] != undefined && rows[nextKey]._cells[0]._value.model.value == '$$fpo') {
            break; // Stop the loop when $$fpo is found
        }

        for (let i = 1; i <= 50; i++) {
            // Check if the cell is present and update its value
            if (!rows[nextKey]._cells[i]) {
                updateCellValue(nextKey, i, 'dummy');
            } else {
                updateCellValue(nextKey, i, ' ');
            }
        }

        nextKey++;
    }
}

                // let add1 = rows[key1]._cells[3]._value.model.address;
                // let add2 = rows[key1]._cells[4]._value.model.address; 
                // let add3 = rows[key1]._cells[5]._value.model.address;
                // data = { ...data, [add1]: 'β' };
                // data = { ...data, [add2]: '=' };
                // data = { ...data, [add3]: beta1.toFixed(3) };

                // let add4 = rows[key1]._cells[8]._value.model.address;
                // let add5 = rows[key1]._cells[9]._value.model.address;
                // let add6 = rows[key1]._cells[10]._value.model.address;
                // data = { ...data, [add4]: 'θ' };
                // data = { ...data, [add5]: '=' };
                // data = { ...data, [add6]: theta1.toFixed(3) };

                // let add7 = rows[key1]._cells[13]._value.model.address;
                // let add8 = rows[key1]._cells[14]._value.model.address;
                // let add9 = rows[key1]._cells[15]._value.model.address;
                // data = { ...data, [add7]:  'εₓ'  };
                // data = { ...data, [add8]: '=' };
                // data = { ...data, [add9]: Exm.toFixed(8) };
            
            if (rows[key1]._cells[0] != undefined && rows[key1]._cells[0]._value.model.value == '$$A') {
                console.log(rows[key1]._cells);
                let cell = rows[key1]._cells[4];

                // Check if cell and its properties are defined
                if (cell && cell._address) {
                    let add11 = cell._address;
                    data = { ...data, [add11]: 'A' };
                } else {
                    // Handle the case where _address is undefined or not available
                    console.error("Error: Unable to determine address for rows[key1]._cells[4]");  
                    // Handle this error scenario appropriately
                }
                let cell2 = rows[key1]._cells[6];
                
                // Check if cell and its properties are defined
                if (cell2 && cell2._address) {
                    let add12 = cell2._address;
                    data = { ...data, [add12]: 'Aₘᵢₙ' };
                } else {
                    // Handle the case where _address is undefined or not available
                    console.error("Error: Unable to determine address for rows[key1]._cells[4]");
                    // Handle this error scenario appropriately
                }
                let cell3 = rows[key1]._cells[5];
                let comparisonSymbol = Av >= Avm ? '≥' : '<';
                if (cell3 && cell3._address) {
                    let add13 = cell3._address;
                    data = { ...data, [add13]: comparisonSymbol }
                }
            }
            if (rows[key1]._cells[0] != undefined && rows[key1]._cells[0]._value.model.value == '$$e') {
                console.log(rows[key1]._cells);
                let cell = rows[key1]._cells[4]; 
                if (cell && cell._address) {
                    let add11 = cell._address;
                    data = { ...data, [add11]: 'Ex' };
                } else {
                    // Handle the case where _address is undefined or not available
                    console.error("Error: Unable to determine address for rows[key1]._cells[4]");
                    // Handle this error scenario appropriately
                }
                let cell2 = rows[key1]._cells[5];
                
                // Check if cell and its properties are defined
                if (cell2 && cell2._address) {
                    let add12 = cell2._address;
                    data = { ...data, [add12]: '=' };
                } else {
                    // Handle the case where _address is undefined or not available
                    console.error("Error: Unable to determine address for rows[key1]._cells[4]");
                    // Handle this error scenario appropriately
                }
                let cell3 = rows[key1]._cells[6];
                if (cell3 && cell3._address){
                    let add13 = cell3._address;
                    let cell3Value = Av >= Avm ? Exm : Etm;
                    data = { ...data, [add13]: cell3Value };
                }
            }
            if (rows[key1]._cells[0] != undefined && rows[key1]._cells[0]._value.model.value == '$$sx') {
                rows[key1]._cells = rows[key1]._cells.map(cell => cell === "" ? undefined : cell);
                console.log(rows[key1]._cells);
                let cell = rows[key1]._cells[4]; 
                if (cell && cell._address) {
                    let add11 = cell._address;
                    data = { ...data, [add11]: 'sx' };
                } else {
                    // Handle the case where _address is undefined or not available
                    console.error("Error: Unable to determine address for rows[key1]._cells[4]");
                    // Handle this error scenario appropriately
                }
                let cell2 = rows[key1]._cells[5];
                
                // Check if cell and its properties are defined
                if (cell2 && cell2._address) {
                    let add12 = cell2._address;
                    data = { ...data, [add12]: '=' };
                } else {
                    // Handle the case where _address is undefined or not available
                    console.error("Error: Unable to determine address for rows[key1]._cells[5]");
                    // Handle this error scenario appropriately
                }
                let cell3 = rows[key1]._cells[33];
                if (cell3 && cell3._address) {
                    let add13 = cell3._address;
                    // let add15value = dv < sg ? dv : sg;
                    data = { ...data, [add13]: `Min| dv, maximum distance between the longitudinal r/f |  =  ${add15value}` };              
                } else {
                    // Handle the case where _address is undefined or not available
                    console.error("Error: Unable to determine address for rows[key1]._cells[33]")
                }
                 let cell4 = rows[key1]._cells[26];
                if (cell4 && cell4._address) {
                    let add14 = cell4._address;
                    data = { ...data, [add14]: '=' };
                } else {
                    // Handle the case where _address is undefined or not available
                    console.error("Error: Unable to determine address for rows[key1]._cells[24]");
                    // Handle this error scenario appropriately
                }
                let cell5 = rows[key1]._cells[25];
                if (cell5 && cell5._address) {
                    let add15 = cell5._address;
                    let add15value = dv < sg ? dv : sg;
                    data = { ...data, [add15]: add15value};
                }else {
                    // Handle the case where _address is undefined or not available
                    console.error("Error: Unable to determine address for rows[key1]._cells[25]");
                    // Handle this error scenario appropriately
                }
            }
            if (rows[key1]._cells[0] != undefined && rows[key1]._cells[0]._value.model.value == '$$sxe') {
                let cell3 = rows[key1]._cells[4];
                if (cell3 && cell3._address) {
                 let add13 = cell3._address;
                 // let sxe = ((add15value*1.38)/(ag +0.63))
                 data ={ ...data, [add13]: 'sxe'}
                }
                let cell4 = rows[key1]._cells[5];
                if (cell4 && cell4._address) {
                 let add14 = cell4._address;
                 // let sxe = ((add15value*1.38)/(ag +0.63))
                 data ={ ...data, [add14]: '='}
                }
                let cell5 = rows[key1]._cells[6];
                if (cell5 && cell5._address) {
                 let add15 = cell5._address;
                 
                 data ={ ...data, [add15]: sxe}
                }               
         }
         if (rows[key1]._cells[0] != undefined && rows[key1]._cells[0]._value.model.value == '$$b') {
            if (Av >= Avm) {
                let cell7 = rows[key1]._cells[7];
                if (cell7 && cell7._address) {
                    let add17 = cell7._address;
                    data = { ...data, [add17]: beta1 };
                } else {
                    // Handle the case where _address is undefined or not available
                    console.error("Error: Unable to determine address for rows[key1]._cells[7]");
                    // Handle this error scenario appropriately
                }
            }

         }
         if (rows[key1]._cells[0] != undefined && rows[key1]._cells[0]._value.model.value == '$$theta') {
            if (Av >= Avm) {
                let cell7 = rows[key1]._cells[7];
                if (cell7 && cell7._address) {
                    let add17 = cell7._address;
                    data = { ...data, [add17]: theta1 };
                } else {
                    // Handle the case where _address is undefined or not available
                    console.error("Error: Unable to determine address for rows[key1]._cells[7]");
                    // Handle this error scenario appropriately
                }
            }

         }
       
         if (rows[key1]._cells[0] != undefined && rows[key1]._cells[0]._value.model.value == '$$beta') {
            let cell9 = rows[key1]._cells[9];
            if (cell9 && cell9._address) {
                beta = cell9._value.model.value;
            } else {
                // Handle the case where _address is undefined or not available
                console.error("Error: Unable to determine address for rows[key1]._cells[9]");
                // Handle this error scenario appropriately
            }
         }
         if (rows[key1]._cells[0] != undefined && rows[key1]._cells[0]._value.model.value == '$$vc') {
            let cell13 = rows[key1]._cells[13];
            if (cell13 && cell13._address) {
                // Retrieve the value from cell13, divide it by beta, and multiply by beta1
                let value13 = cell13._value.model.value;
                let result = (value13 / beta) * beta1;
        
                // Store the result back in cell13 or handle it as needed
                cell13._value.model.value = result;
            } else {
                // Handle the case where _address is undefined or not available
                console.error("Error: Unable to determine address for rows[key1]._cells[13]");
                // Handle this error scenario appropriately
            }
         }

            if (rows[key1]._cells[0] != undefined && rows[key1]._cells[0]._value.model.value == '$$strn') {
                let add1 = rows[key1]._cells[2]._value.model.address;
                let add2 = rows[key1]._cells[8]._value.model.address;
                data = { ...data, [add1]: 'Calculation for β and θ' };
                data = { ...data, [add2]: '' };
            }
            if (rows[key1]._cells[0] != undefined && rows[key1]._cells[0]._value.model.value == '$$strn1') {
                let add1 = rows[key1]._cells[3]._value.model.address;
                let add2 = rows[key1]._cells[4]._value.model.address; let add3 = rows[key1]._cells[5]._value.model.address;
                data = { ...data, [add1]: 'β' };
                data = { ...data, [add2]: '=' };
                data = { ...data, [add3]: beta2.toFixed(3) };

                let add4 = rows[key1]._cells[8]._value.model.address;
                let add5 = rows[key1]._cells[9]._value.model.address;
                let add6 = rows[key1]._cells[10]._value.model.address;
                data = { ...data, [add4]: 'θ' };
                data = { ...data, [add5]: '=' };
                data = { ...data, [add6]: theta2.toFixed(3) }; 

                let add7 = rows[key1]._cells[13]._value.model.address;
                let add8 = rows[key1]._cells[14]._value.model.address;
                let add9 = rows[key1]._cells[15]._value.model.address;
                data = { ...data, [add7]: 'εₓ' };
                data = { ...data, [add8]: '=' };
                data = { ...data, [add9]: Exn.toFixed(8) };
            }
            if (rows[key1]._cells[0] != undefined && rows[key1]._cells[0]._value.model.value == '$$Vc') {
                let add1 = rows[key1]._cells[9]._value.model.address;
                data = { ...data, [add1]: Vc1 };
            }
        }
        for (let key in data) {
            let match =  key.match(/^([A-Za-z]+)(\d+)$/);
            if (match) {
                const row = match[1];
                const col = match[2];
                let value = 0;
                let factor = 1;
                for (let i = row.length - 1; i >= 0; i--) {
                    value += (row.charCodeAt(i) - 64) * factor;
                    factor *= 26;
                }
                worksheet._rows[col - 1]._cells[value - 1]._value.model.value = data[key];
                // worksheet={...worksheet.worksheet._rows[col - 1].cells,[value-1]._value.model.value:data[key]}
                worksheet._rows[col - 1]._cells[value - 1]._value.model.type = 3;
                if (data[key] == 'β') {
                    for (let i = col; i <= Number(col) + 5; i++) {
                        // console.log(i,col);
                        delete worksheet._rows[i];
                    }
                }
            }
        }
        workbookData.worksheets[wkey] = worksheet;
        setWorkbookData(workbookData);
        // setWorkbookData( { ...workbookData.worksheet[wkey], [wkey]: worksheet});
       
        // console.log(worksheet);
        setSheetName(worksheet.name);
    }
    async function fetchLc() {
        const endpointsDataKeys = [
            { endpoint: "/db/lcom-gen", dataKey: "LCOM-GEN" },
            { endpoint: "/db/lcom-conc", dataKey: "LCOM-CONC" },
            // { endpoint: "/db/lcom-src", dataKey: "LCOM-SRC" },
            // { endpoint: "/db/lcom-steel", dataKey: "LCOM-STEEL" },
            // { endpoint: "/db/lcom-stlcomp", dataKey: "LCOM-STLCOMP" },
        ];
        let check = false;
        let lc;
        try {
            for (const { endpoint, dataKey } of endpointsDataKeys) {
                const response = await midasAPI("GET", endpoint);
                if (response && !response.error) {
                    setLc(response[dataKey]);
                    lc = response[dataKey];
                    console.log(response[dataKey])
                    check = true;
                }
            }

            if (!check) {
                enqueueSnackbar("Please Check Connection And Defined Load Combination", {
                    variant: "error",
                    anchorOrigin: {
                        vertical: "top",
                        horizontal: "center"
                    },
                });
                setLc([' ']);
                return null;
            }
            showLc(lc);
        } catch (error) {
            // console.error(`Error fetching data from ${endpoint}:`, error);
            enqueueSnackbar("Unable to Fetch Data Check Connection", {
                variant: "error",
                anchorOrigin: {
                    vertical: "top",
                    horizontal: "center"
                },
            });
            return null;
        }

    } // End fetching load combination

    function showLc(lc) {
        console.log(lc);
        item.delete('1');
        for (let key in lc) {
            items.set(lc[key].NAME, key)
            // console.log(key, lc[key].NAME)
        }

        setItem(items);
    }

    const handleFileDownload = async () => {
        // fetchLc();
        const combArray = Object.values(Lc);
        if (combArray.length === 0) {
            enqueueSnackbar("Please Define Load Combination", {
                variant: "error",
                anchorOrigin: {
                    vertical: "top",
                    horizontal: "center"
                },
            });
            return;
        }
        // console.log('load combinations', Lc)
        console.log(SelectWorksheets);

        for (let wkey in SelectWorksheets) {
            updatedata(wkey, SelectWorksheets[wkey]);
        }

        if (!workbookData) return;
        const worksheet = workbookData.getWorksheet(sheetName);
        const buffer = await workbookData.xlsx.writeBuffer();
        const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });

        const link = document.createElement('a');
        link.href = URL.createObjectURL(blob);
        link.download = 'output.xlsx';
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
            <div >
                <Typography variant="h1"  > Casting Method</Typography>
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
            </div >

            <div style={{ marginTop: "25px" }}>

                <Grid container>
                    <Grid item xs={9}>
                        <Typography variant="h1"  >  Maximum Spacing of Transverse Reinforcement:</Typography>
                    </Grid>
                    <Grid item xs={3}>
                        <Typography variant="h1"  > (5.7.2.6.-1)</Typography>
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
                        <Radio
                            name="CA (18 inches)"
                            value="ca1"
                            checked={sp === "ca1"}
                        />
                        <Radio
                            name="AASHTO LFRD (24 inches)"
                            value="aa1"
                            checked={sp === "aa1"}
                        />
                    </div>
                </RadioGroup>
            </div >

            <div style={{ marginTop: "25px" }}>

                <Grid container>
                    <Grid item xs={9}>
                        <Typography variant="h1"  > Clear Concrete Cover:</Typography>
                    </Grid>
                    <Grid item xs={3}>
                        <Typography variant="h1"  > (5.6.7-1)</Typography>
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
                        <Radio
                            name="CA (2.5 inches)"
                            value="ca2"
                            checked={cvr === "ca2"}
                        />
                        <Radio
                            name="AASHTO LFRD (1.8 inches)"
                            value="aa2"
                            checked={cvr === "aa2"}
                        />
                    </div>
                </RadioGroup>
            </div >
            <div style={{ marginTop: "25px" }}>
                <Grid container>
                    <Grid item xs={3}>
                        <Typography variant="h1"  > Load Case for SLS (Permanent Loads)</Typography>
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
                        <Typography variant="h1"  >(5.9.2.3.2b-1)</Typography>
                    </Grid>
                </Grid>
            </div>

            <div
                style={{ display: "flex", flexDirection: "column", justifyContent: "space-between", marginTop: "20px" }}
            >
                <Grid container>
                    <Grid item xs={6}>
                        <Typography variant="h1" height="40px" paddingTop="15px"  >
                            <input type="file" accept=".xlsx, .xls" onChange={handleFileUpload} />
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
 <Grid container direction='row'>
    <Grid
      item
      xs={6}
    >
        <Typography>
  Maximum aggregate size(ag)
</Typography>
</Grid>
<Grid
      item
      xs={6}
    >
      <TextField
  value={ag}
  onChange={handleAgChange}
  placeholder=""
//   title="Maximum aggregate size(ag)"
  width="100px"
/>
</Grid>
<Grid
      item
      xs={6}
      marginTop={0.5}
    >
        <Typography size='small'>
  Maximum distance between the longitudinal reinforcement
</Typography>
</Grid>
    <Grid
      item
      xs={6}
      marginTop={0.5}
    >
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

}
