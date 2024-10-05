import * as XLSX from 'xlsx';
import './App.css';

const jsonData = [];
const getExcel = () => {
    const worksheet = XLSX.utils.json_to_sheet(jsonData);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, 'Sheet1');
    XLSX.writeFile(workbook, 'output.xlsx');
};

const storeAllRequests = event => {
    const allRequests = event.target.value;
    if (allRequests) {
        const data = allRequests
            .replaceAll('*', '')
            .split('\n')
            .filter(a => a && a.includes(':'));
        let dict;
        for (const s of data) {
            if (s.includes('WARDHA JN')) {
                if (dict) {
                    jsonData.push(dict);
                }
                dict = {};
            } else {
                const [col, val] = s.split(':');
                dict[col] = val;
            }
        }
    }
};
const App = () => {
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
                        onInput={storeAllRequests}
                    ></textarea>
                </div>
                <button name="Export to Excel" onClick={getExcel}>
                    {'Export to Excel'}
                </button>
            </header>
        </div>
    );
};

export default App;
