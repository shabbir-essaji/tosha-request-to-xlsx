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
        const jsonData = getJsonData();
        const worksheet = XLSX.utils.json_to_sheet(jsonData);
        util.autofitColumns(worksheet);
        const merge = [
            { s: { r: 1, c: 0 }, e: { r: 2, c: 0 } },
            { s: { r: 3, c: 0 }, e: { r: 4, c: 0 } }
        ];
        worksheet['!merges'] = merge;
        const workbook = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(workbook, worksheet, 'Sheet1');
        XLSX.writeFile(
            workbook,
            `Tosha Requests for ${moment(dateRef.current, 'DD/MM/YYYY').format('DD MMM YYYY')}.xlsx`
        );
        setTextVal('');
        setTextAreaDisabled(false);
    };

    const getJsonData = () => {
        const arrayOfObjects = [];
        if (textVal) {
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
        }
        return arrayOfObjects;
    };
    return (
        <div className="App">
            <header className="App-header">
                <div style={{ display: 'flex', flexDirection: 'column' }}>
                    <label htmlFor="toshaRequests">
                        Paste all tosha requests in 1 go
                    </label>
                    <textarea
                        id="toshaRequests"
                        name="toshaRequests"
                        rows={40}
                        cols={50}
                        val={textVal}
                        onInput={event => setTextVal(event.target.value)}
                        disabled={textAreaDisabled}
                    ></textarea>
                </div>
                <button
                    name="Export to Excel"
                    onClick={() => {
                        setTextAreaDisabled(true);
                        getExcel();
                    }}
                    disabled={!textVal}
                >
                    {'Export to Excel'}
                </button>
            </header>
        </div>
    );
};

export default App;
