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
            /*
        const merge = [
            { s: { r: 1, c: 0 }, e: { r: 2, c: 0 } },
            { s: { r: 3, c: 0 }, e: { r: 4, c: 0 } }
        ];
        worksheet['!merges'] = merge;
        */
            const workbook = XLSX.utils.book_new();
            XLSX.utils.book_append_sheet(workbook, worksheet, 'Sheet1');
            XLSX.writeFile(
                workbook,
                `Tosha Requests for ${moment(dateRef.current, 'DD/MM/YYYY').format('DD MMM YYYY')}.xlsx`
            );
        } catch (e) {
            alert(
                'Unable to parse input. Kindly recheck the input tosha input'
            );
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
                .filter(a => a && a.includes(':'));
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
                            obj[col] = extra
                                ? `${val}:${extra}`.trim()
                                : val.trim();
                        } else if (('Date' === col) & !dateRef.current) {
                            dateRef.current = val.trim();
                        }
                        from++;
                    }
                    arrayOfObjects.push(obj);
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
