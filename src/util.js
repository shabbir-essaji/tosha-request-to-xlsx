import { range } from 'lodash';

const getEndingDigit = inputString => {
    // Define a regular expression to extract the ending digits
    const regex = /\d+$/;

    // Use the test method to check if the inputString matches the pattern
    if (regex.test(inputString)) {
        // Use the match method to extract the ending digits from the inputString
        const match = inputString.match(regex);
        // Remove any non-digit characters before returning the result
        return match[0].replace(/\D/g, '');
    }

    // Return null if there are no ending digits
    return null;
};
const autofitColumns = worksheet => {
    //using https://github.com/SheetJS/sheetjs/issues/1473#issuecomment-1641573655
    let objectMaxLength = [];

    const [startLetter, endLetter] = worksheet['!ref']
        ?.replace(/\d/, '')
        .split(':');
    const ranges = range(
        startLetter.charCodeAt(0),
        endLetter.charCodeAt(0) + 1
    );

    ranges.forEach(c => {
        const cellHeader = String.fromCharCode(c);

        const maxCellLengthForWholeColumn = Array.from(
            { length: getEndingDigit(worksheet['!ref']) - 1 },
            (_, i) => i
        ).reduce((acc, i) => {
            const cell = worksheet[`${cellHeader}${i + 1}`];

            // empty cell
            if (!cell) return acc;

            const charLength = cell.v.length + 1;

            return acc > charLength ? acc : charLength;
        }, 0);

        objectMaxLength.push({ wch: maxCellLengthForWholeColumn });
    });
    worksheet['!cols'] = objectMaxLength;
};

const util = {
    autofitColumns
};
export default util;
