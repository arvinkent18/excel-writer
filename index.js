const excel = require('exceljs');
const fs = require('fs');

const data = JSON.parse(fs.readFileSync('data.json'));

let workbook = new excel.Workbook();

let worksheet = workbook.addWorksheet('sheet1');

worksheet.columns = [
    { header: 'SLNOINPART', key: 'SLNOINPART' },
    { header: 'C_HOUSE_NO', key: 'C_HOUSE_NO' },
    { header: 'C_HOUSE_NO_V1', key: 'C_HOUSE_NO_V1' },
    { header: 'FM_NAME_EN', key: 'FM_NAME_EN' },
    { header: 'LASTNAME_EN', key: 'LASTNAME_EN' },
    { header: 'FM_NAME_V1', key: 'FM_NAME_V1' },
    { header: 'LASTNAME_V1', key: 'LASTNAME_V1' },
    { header: 'RLN_TYPE', key: 'RLN_TYPE' },
    { header: 'RLN_FM_NM_EN', key: 'RLN_FM_NM_EN' },
    { header: 'RLN_L_NM_EN', key: 'RLN_L_NM_EN' },
    { header: 'RLN_FM_NM_V1', key: 'RLN_FM_NM_V1' },
    { header: 'RLN_L_NM_V1', key: 'RLN_L_NM_V1' },
    { header: 'EPIC_NO', key: 'EPIC_NO' },
    { header: 'GENDER', key: 'GENDER' },
    { header: 'AGE', key: 'AGE' },
]

data.forEach((e, _index) => {
    worksheet.addRow({
        ...e,
    })
})

workbook.xlsx.writeFile('expected_output.xlsx')