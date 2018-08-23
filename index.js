const XLSX = require('xlsx');
const excel = require('node-excel-export');
const fs = require('fs')
const workbook = XLSX.readFile('./list.xlsx', {});

// 获取 Excel 中所有表名
const sheetNames = workbook.SheetNames; // 返回 ['sheet1', 'sheet2']
// 根据表名获取对应某张表
const worksheet = workbook.Sheets[sheetNames[0]];

let sheet1 = XLSX.utils.sheet_to_json(worksheet);

sheet1.forEach(item => {
  var arr = item['清单'].split(/[-_－\s～—–~，𡿨／／/]/)
  arr.forEach((name, index) => {
    item['col' + index] = name
  })
})

// You can define styles as json object
const styles = {
  headerDark: {
    fill: {
      fgColor: {
        rgb: 'FF000000'
      }
    },
    font: {
      color: {
        rgb: 'FFFFFFFF'
      },
      sz: 14,
      bold: true,
      underline: true
    }
  },
  cellPink: {
    fill: {
      fgColor: {
        rgb: 'FFFFCCFF'
      }
    }
  },
  cellGreen: {
    fill: {
      fgColor: {
        rgb: 'FF00FF00'
      }
    }
  }
};

//Array of objects representing heading rows (very top)
const heading = [];

//Here you specify the export structure
const specification = {
  '序号': { // <- the key should match the actual data key
    displayName: '序号', // <- Here you specify the column header
    headerStyle: styles.headerDark, // <- Header style
    cellStyle: styles.cellGreen,
    width: 50 // <- width in pixels
  },
  '清单': {
    displayName: '清单',
    headerStyle: styles.headerDark,
    // cellFormat: function (value, row) { // <- Renderer function, you can access also any row.property
    //   return (value == 1) ? 'Active' : 'Inactive';
    // },
    width: 400 // <- width in chars (when the number is passed as string)
  },
  '所属微信群': {
    displayName: '所属微信群',
    headerStyle: styles.headerDark,
    cellStyle: styles.cellPink, // <- Cell style
    width: 120 // <- width in pixels
  },
  'col0': {
    displayName: 'col0',
    headerStyle: styles.headerDark,
    cellStyle: styles.cellPink, // <- Cell style
    cellFormat: function (value, row) { // <- Renderer function, you can access also any row.property
      return value ? value : '';
    },
    width: 200 // <- width in pixels
  },
  'col1': {
    displayName: 'col1',
    headerStyle: styles.headerDark,
    cellStyle: styles.cellPink, // <- Cell style
    cellFormat: function (value, row) { // <- Renderer function, you can access also any row.property
      return value ? value : '';
    },
    width: 120 // <- width in pixels
  },
  'col2': {
    displayName: 'col2',
    headerStyle: styles.headerDark,
    cellStyle: styles.cellPink, // <- Cell style
    cellFormat: function (value, row) { // <- Renderer function, you can access also any row.property
      return value ? value : '';
    },
    width: 120 // <- width in pixels
  },
  'col3': {
    displayName: 'col3',
    headerStyle: styles.headerDark,
    cellStyle: styles.cellPink, // <- Cell style
    cellFormat: function (value, row) { // <- Renderer function, you can access also any row.property
      return value ? value : '';
    },
    width: 120 // <- width in pixels
  },
  'col4': {
    displayName: 'col4',
    headerStyle: styles.headerDark,
    cellStyle: styles.cellPink, // <- Cell style
    cellFormat: function (value, row) { // <- Renderer function, you can access also any row.property
      return value ? value : '';
    },
    width: 120 // <- width in pixels
  },
  'col5': {
    displayName: 'col5',
    headerStyle: styles.headerDark,
    cellStyle: styles.cellPink, // <- Cell style
    cellFormat: function (value, row) { // <- Renderer function, you can access also any row.property
      return value ? value : '';
    },
    width: 120 // <- width in pixels
  }
}

// Define an array of merges. 1-1 = A:1
// The merges are independent of the data.
// A merge will overwrite all data _not_ in the top-left cell.
const merges = []

// Create the excel report.
// This function will return Buffer
const report = excel.buildExport(
  [ // <- Notice that this is an array. Pass multiple sheets to create multi sheet report
    {
      name: 'Report', // <- Specify sheet name (optional)
      heading: heading, // <- Raw heading array (optional)
      merges: merges, // <- Merge cell ranges
      specification: specification, // <- Report specification
      data: sheet1 // <-- Report data
    }
  ]
);

fs.writeFile('./list_new.xlsx', report, (err) => {
  if (err) throw err;
  console.log('文件已保存！');
});