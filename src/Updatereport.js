import { DropList, Grid, Panel, Typography, VerifyUtil } from '@midasit-dev/moaui';
import { Radio, RadioGroup } from "@midasit-dev/moaui";
import React, { useState } from 'react';
import * as Buttons from "./Components/Buttons";
import ExcelJS from 'exceljs';
import AlertDialogModal from './AlertDialogModal';
import { midasAPI } from "./Function/Common";
import { enqueueSnackbar } from 'notistack';
import { ThetaBeta1 } from './Function/ThetaBeta';


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
    let names = {};
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
        CalBetaTheta();
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

            if (rows[key1]._cells[0] != undefined && rows[key1]._cells[0]._value.model.value == '$$dc') {
                if (cvr === "ca2") {
                    let add1 = rows[key1]._cells[8]._value.model.address;

                    let add2 = rows[key1]._cells[11]._value.model.address;
                    let val2 = rows[key1]._cells[11]._value.model.value + 3.6 - 5;
                    data = { ...data, [add1]: '2*2.5' };
                    data = { ...data, [add2]: val2 };
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
            }
        }
        console.log(Mmax, Mmin, Sb, St, Nmax, Nmin, Ag, E,worksheet)
        // console.log(worksheet);
        for (let key in data) {
            const match = key.match(/^([A-Za-z]+)(\d+)$/);
            if (match) {
                const row = match[1];
                const col = match[2];
                // console.log("Letter Part:", row); // Output: "AD"
                // console.log("Number Part:", col); // Output: "111"
                let value = 0;
                let factor = 1;
                for (let i = row.length - 1; i >= 0; i--) {
                    value += (row.charCodeAt(i) - 64) * factor;
                    factor *= 26;
                }

                // console.log(col - 1,value - 1)
                worksheet._rows[col - 1]._cells[value - 1]._value.model.value = data[key];
            }
        }
        workbookData.worksheets[wkey] = worksheet;
        setWorkbookData(workbookData);
        // setSheetData(jsonData);
        setSheetName(worksheet.name);
    }
    // console.log(workbookData)
    function CalBetaTheta() {
        let value = ThetaBeta1(0.12, 0.078);
        let theta = ThetaBeta1[0];
        let beta = ThetaBeta1[1];
    }

    // to get all the loadcombinations
    async function fetchLc() {
        const endpointsDataKeys = [
            { endpoint: "/db/lcom-gen", dataKey: "LCOM-GEN" },
            { endpoint: "/db/lcom-conc", dataKey: "LCOM-CONC" },
            { endpoint: "/db/lcom-src", dataKey: "LCOM-SRC" },
            { endpoint: "/db/lcom-steel", dataKey: "LCOM-STEEL" },
            { endpoint: "/db/lcom-stlcomp", dataKey: "LCOM-STLCOMP" },
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
        console.log('load combinations', Lc)
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
        <Panel width={520} height={400} marginTop={3} padding={2} variant="shadow2">
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
                style={{ display: "flex", justifyContent: "space-between", marginTop: "35px" }}
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
            </div>


            <div
                style={{
                    display: "flex",
                    justifyContent: "space-between",
                    margin: "0px",
                    marginTop: "30px",
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

