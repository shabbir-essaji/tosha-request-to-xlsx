import { useState, useRef } from 'react';
import * as XLSX from 'xlsx';
import moment from 'moment';
import util from './util.js';
import './App.css';

const App = () => {
    const FILTERED_COLUMNS = ['Day', 'Date', 'ITS ID', 'Email'];
    const [textVal, setTextVal] = useState('');
    const [textAreaDisabled, setTextAreaDisabled] = useState(false);
    const dateRef = useRef('');
    const getExcel = () => {
        try {
            const jsonData = getJsonData();
            const worksheet = XLSX.utils.json_to_sheet(jsonData);
            util.autofitColumns(worksheet);

            const merge = util.getMergedColumnsInfo(worksheet);
            worksheet['!merges'] = merge;

            const workbook = XLSX.utils.book_new();
            XLSX.utils.book_append_sheet(workbook, worksheet, 'Sheet1');
            XLSX.writeFile(
                workbook,
                `Tosha Requests for ${moment(dateRef.current, 'DD/MM/YYYY').format('DD MMM YYYY')}.xlsx`
            );
            setTextAreaDisabled(true);
        } catch (e) {
            alert('Unable to parse input. Kindly recheck the tosha requests');
            setTextAreaDisabled(false);
        } finally {
            setTextVal('');
        }
    };

    const getJsonData = () => {
        if (textVal) {
            const arrayOfObjects = [];
            const data = textVal
                .replaceAll('*', '')
                .split('\n')
                .filter(a => a?.includes(':'));
            let serialNo = 1;
            for (let i = 0; i < data.length; ) {
                if (data[i].includes('WARDHA JN')) {
                    let from = i + 1;
                    const obj = {};
                    while (
                        from < data.length &&
                        !data[from].includes('WARDHA JN')
                    ) {
                        const [col, val, extra] = data[from].split(':');
                        if (!FILTERED_COLUMNS.includes(col)) {
                            if (col === 'No of pax') {
                                obj[col] = Number(val.trim());
                            } else {
                                obj[col] = extra
                                    ? `${val}:${extra}`.trim()
                                    : val.trim();
                            }
                        } else if (('Date' === col) & !dateRef.current) {
                            dateRef.current = val.trim();
                        }
                        from++;
                    }
                    const orderedObj = {
                        'Sr. No.': serialNo++,
                        'Train details': null,
                        Time: null,
                        'No of pax': null
                    };

                    const finalObj = Object.assign(orderedObj, obj);
                    arrayOfObjects.push(finalObj);
                    i = from;
                }
            }
            return arrayOfObjects;
        }
    };
    return (
        <div className="App">
            <h1>Tosha Requests to XLSX</h1>
            <div>
                <textarea
                    id="toshaRequests"
                    name="toshaRequests"
                    placeholder={'Paste all tosha requests here from Whatsapp'}
                    value={textVal}
                    rows={15}
                    cols={60}
                    onInput={event => setTextVal(event.target.value)}
                    disabled={textAreaDisabled}
                />
            </div>
            <button
                name="Export to Excel"
                className={textVal ? 'exportToXlsx' : ''}
                onClick={() => {
                    setTextAreaDisabled(true);
                    getExcel();
                }}
                disabled={!textVal}
            >
                {'Export to Excel'}
            </button>
        </div>
    );
};

export default App;
