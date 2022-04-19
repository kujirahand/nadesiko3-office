/**
 * なでしこ3 プラグイン
 * nadesiko3-office
 * Excelのためのプラグイン
 */

const Excel = require('exceljs')
const ERR_NO_WORKBOOK = 'Excel関連の命令を使う時は、最初に『EXCEL新規ブック』や『EXCEL開』などでブックを用意してください。'

const PluginOffice = {
    '初期化': {
        type: 'func',
        josi: [],
        fn: function (sys) {
            // ここにプラグインの初期化処理
            sys.__varslist[0]['OFFICEバージョン'] = '0.0.1'
            // Excelのインスタンス
            sys.__workbook = null
            sys.__worksheet = null
        }
    },

    // @OFFICE定数
    'OFFICEバージョン': {type: 'const', value:'?'}, // @OFFICEばーじょん

    // @エクセル(Excel)
    'エクセル新規ブック': { // @Excelの新規ワークブックを生成してオブジェクトを返す // @えくせるしんきぶっく
        type: 'func',
        josi: [],
        fn: function (sys) {
            const workbook = new ExcelJS.Workbook()
            sys.__workbook = workbook
            return workbook
        }
    },
    'エクセル開': { // @ファイルFILEからExcelワークブックを読んで返す // @えくせるひらく
        type: 'func',
        josi: [['を', 'の', 'から']],
        asyncFn: true,
        fn: async function (file, sys) {
            const workbook = new Excel.Workbook()
            sys.__workbook = workbook
            await workbook.xlsx.readFile(file)
            if (workbook.worksheets.length > 0) {
                sys.__worksheet = workbook.worksheets[0]
            }
            return workbook
        }
    },
    'エクセル保存': { // @ファイルFILEへ作業中のExcelワークブックを保存する // @えくせるほぞん
        type: 'func',
        josi: [['へ', 'に']],
        asyncFn: true,
        fn: async function (file, sys) {
            if (sys.__workbook === null) {
                throw new Error(ERR_NO_WORKBOOK)
            }
            sys.__workbook.xlsx.writeFile(file)
        },
        return_none: true
    },
    'エクセルCSV保存': { // @ファイルFILEへ作業中のExcelワークブックをCSVで保存する // @えくせるCSVほぞん
        type: 'func',
        josi: [['へ', 'に']],
        asyncFn: true,
        fn: async function (file, sys) {
            if (sys.__workbook === null) {
                throw new Error(ERR_NO_WORKBOOK)
            }
            sys.__workbook.csv.writeFile(file)
        },
        return_none: true
    },
    'エクセル新規シート': { // @Excelの作業中のワークブックに新規シートNAMEを追加して返す // @えくせるしんきしーと
        type: 'func',
        josi: [['の', 'で']],
        fn: function (name, sys) {
            if (sys.__workbook === null) {
                sys.__exec('EXCEL新規ブック', [sys])
            }
            const sheet = sys.__workbook.addWorksheet(name)
            sys.__worksheet = sheet
            return sheet
        }
    },
    'エクセルシート取得': { // @NAMEのシートを取得して返す // @えくせるしーとしゅとく
        type: 'func',
        josi: [['の']],
        fn: function (name, sys) {
            if (sys.__workbook === null) {
                throw new Error(ERR_NO_WORKBOOK)
            }
            const sheet = sys.__workbook.getWorksheet(name)
            return sheet
        }
    },
    'エクセルシート注目': { // @NAMEのシートを取得して返す // @えくせるしーとちゅうもく
        type: 'func',
        josi: [['の','に','を']],
        fn: function (name, sys) {
            if (sys.__workbook === null) {
                throw new Error(ERR_NO_WORKBOOK)
            }
            const sheet = sys.__workbook.getWorksheet(name)
            sys.__worksheet = sheet
            return sheet
        }
    },
    'エクセルセル設定': { // @セルA(例えば「A1」)へVを設定する // @えくせるせるせってい
        type: 'func',
        josi: [['へ','に'],['を']],
        fn: function (cell, v, sys) {
            if (sys.__workbook === null) {
                throw new Error(ERR_NO_WORKBOOK)
            }
            const objCell = sys.__worksheet.getCell(cell)
            objCell.value = v
        },
        return_none: true
    },
    'エクセルセル取得': { // @セルA(例えば「A1」)の値を取得して返す // @えくせるせるしゅとく
        type: 'func',
        josi: [['から','を','の']],
        fn: function (cell, sys) {
            if (sys.__workbook === null) {
                throw new Error(ERR_NO_WORKBOOK)
            }
            const objCell = sys.__worksheet.getCell(cell)
            return objCell.value
        }
    },
    'エクセルシート一覧取得': { // @作業中のブックのシート一覧取得して返す // @えくせるしーといちらんしゅとく
        type: 'func',
        josi: [],
        fn: function (sys) {
            if (sys.__workbook === null) {
                throw new Error(ERR_NO_WORKBOOK)
            }
            const names = []
            sys.__workbook.eachSheet(function(sheet, id){
                names.push(sheet.name)
            })
            return names
        }
    },
    'エクセルシート削除': { // @作業中のブックのシートNAMEを削除する // @えくせるしーとさくじょ
        type: 'func',
        josi: [['の','を']],
        fn: function (name, sys) {
            if (sys.__workbook === null) {
                throw new Error(ERR_NO_WORKBOOK)
            }
            let sheet = sys.__workbook.getWorksheet(name)
            if (!sheet) { throw new Error(`『EXCELシート削除』でシート『${name}』が見当たりません。`) }
            sys.__workbook.removeWorksheet(sheet.id)
        },
        return_none: true
    },
    'エクセルシート削除': { // @作業中のブックのシートNAMEを削除する // @えくせるしーとさくじょ
        type: 'func',
        josi: [['の','を']],
        fn: function (name, sys) {
            if (sys.__workbook === null) {
                throw new Error(ERR_NO_WORKBOOK)
            }
            let sheet = sys.__workbook.getWorksheet(name)
            if (!sheet) { throw new Error(`『EXCELシート削除』でシート『${name}』が見当たりません。`) }
            sys.__workbook.removeWorksheet(sheet.id)
        },
        return_none: true
    },
}

// モジュールのエクスポート(必ず必要)
// scriptタグで取り込んだ時、自動で登録する
if (typeof (navigator) === 'object' && typeof (navigator.nako3) === 'object') {
    navigator.nako3.addPluginObject('PluginOffice', PluginOffice)
} else {
    module.exports = PluginOffice
}


