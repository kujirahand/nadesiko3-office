/**
 * なでしこ3 プラグイン
 * nadesiko3-office
 * Excelのためのプラグイン
 */

const Excel = require('exceljs')
const Color = require('./colors.js')
const Utils = require('./utils.js')
const ERR_NO_WORKBOOK = 'Excel関連の命令を使う時は、最初に『EXCEL新規ブック』や『EXCEL開』などでブックを用意してください。'

const PluginOffice = {
    'meta': {
        type: 'const',
        value: { // プラグインに関する情報を指定する
            pluginName: 'nadesiko3-office', // プラグインの名前
            description: 'Excelプラグイン', // プラグインの説明
            pluginVersion: '3.6.16', // プラグインのバージョン
            nakoRuntime: ['cnako'], // 対象ランタイム
            nakoVersion: '3.6.16' // 要求なでしこバージョン
        }
    },
    '初期化': {
        type: 'func',
        josi: [],
        fn: function (sys) {
            // ここにプラグインの初期化処理
            sys.__setSysVar('OFFICEバージョン', '3.6.16')
            // Excelのインスタンス
            sys.tags.__workbook = null
            sys.tags.__worksheet = null
        }
    },

    // @OFFICE定数
    'OFFICEバージョン': {type: 'const', value:'?'}, // @OFFICEばーじょん

    // @エクセル(Excel)
    'エクセル新規ブック': { // @Excelの新規ワークブックを生成してオブジェクトを返す // @えくせるしんきぶっく
        type: 'func',
        josi: [],
        fn: function (sys) {
            const workbook = new Excel.Workbook()
            sys.tags.__workbook = workbook
            sys.tags.__worksheet = workbook.addWorksheet()
            return workbook
        }
    },
    'エクセル開': { // @ファイルFILEからExcelワークブックを読んで返す // @えくせるひらく
        type: 'func',
        josi: [['を', 'の', 'から']],
        asyncFn: true,
        fn: async function (file, sys) {
            const workbook = new Excel.Workbook()
            sys.tags.__workbook = workbook
            await workbook.xlsx.readFile(file)
            if (workbook.worksheets.length > 0) {
                sys.tags.__worksheet = workbook.worksheets[0]
            }
            return workbook
        }
    },
    'エクセル保存': { // @ファイルFILEへ作業中のExcelワークブックを保存する // @えくせるほぞん
        type: 'func',
        josi: [['へ', 'に']],
        asyncFn: true,
        fn: async function (file, sys) {
            if (sys.tags.__workbook === null) {
                throw new Error(ERR_NO_WORKBOOK)
            }
            sys.tags.__workbook.xlsx.writeFile(file)
        },
        return_none: true
    },
    'エクセルCSV保存': { // @ファイルFILEへ作業中のExcelワークブックをCSVで保存する(ただしUTF-8のCSVとなる) // @えくせるCSVほぞん
        type: 'func',
        josi: [['へ', 'に']],
        asyncFn: true,
        fn: async function (file, sys) {
            if (sys.tags.__workbook === null) {
                throw new Error(ERR_NO_WORKBOOK)
            }
            sys.tags.__workbook.csv.writeFile(file)
        },
        return_none: true
    },
    'エクセル新規シート': { // @Excelの作業中のワークブックに新規シートNAMEを追加して返す // @えくせるしんきしーと
        type: 'func',
        josi: [['の', 'で']],
        fn: function (name, sys) {
            if (sys.tags.__workbook === null) {
                sys.__exec('EXCEL新規ブック', [sys])
            }
            const sheet = sys.tags.__workbook.addWorksheet(name)
            sys.tags.__worksheet = sheet
            return sheet
        }
    },
    'エクセルシート取得': { // @NAMEのシートを取得して返す // @えくせるしーとしゅとく
        type: 'func',
        josi: [['の']],
        fn: function (name, sys) {
            if (sys.tags.__workbook === null) {
                throw new Error(ERR_NO_WORKBOOK)
            }
            const sheet = sys.tags.__workbook.getWorksheet(name)
            return sheet
        }
    },
    'エクセルシート注目': { // @NAMEのシートを取得して返す // @えくせるしーとちゅうもく
        type: 'func',
        josi: [['の','に','を']],
        fn: function (name, sys) {
            if (sys.tags.__workbook === null) {
                throw new Error(ERR_NO_WORKBOOK)
            }
            const sheet = sys.tags.__workbook.getWorksheet(name)
            sys.tags.__worksheet = sheet
            return sheet
        }
    },
    'エクセルセル設定': { // @セル(例えば「A1」)へVを設定する // @えくせるせるせってい
        type: 'func',
        josi: [['へ','に'],['を']],
        fn: function (cell, v, sys) {
            if (sys.tags.__workbook === null) {
                throw new Error(ERR_NO_WORKBOOK)
            }
            const objCell = sys.tags.__worksheet.getCell(cell)
            if (v.substring(0, 1) === '=') {
                objCell.value = { formula: v.substring(1) }
            } else {
                objCell.value = v
            }
        },
        return_none: true
    },
    'エクセル設定': { // @セル(例えば「A1」)へVを設定する // @えくせるせってい
        type: 'func',
        josi: [['へ','に'],['を']],
        fn: function (cell, v, sys) {
            sys.__exec('エクセルセル設定', [cell, v, sys])
        },
        return_none: true
    },
    'エクセル一括設定': { // @左上のセル(例えば「A1」)を起点にして、二次元配列変数VALUESを一括設定する // @えくせるいっかつせってい
        type: 'func',
        josi: [['へ','に'],['を']],
        fn: function (cell, values, sys) {
            if (sys.tags.__workbook === null) {
                throw new Error(ERR_NO_WORKBOOK)
            }
            // LeftTop
            if (cell.indexOf(':') >= 0) {
                cell = cell.split(':')[0]
            }
            // update
            const start = Utils.addressToPos(cell)
            for (let row = 0; row < values.length; row++) {
                const cells = values[row]
                let excelRow = sys.tags.__worksheet.getRow(start.row + row)
                for (let col = 0; col < cells.length; col++) {
                    const val = cells[col]
                    excelRow.getCell(start.col + col).value = val
                }
                sys.tags.__worksheet.getRow(start.row + row).commit()
            }
        },
        return_none: true
    },
    'エクセルセル取得': { // @セル(例えば「A1」)の値を取得して返す // @えくせるせるしゅとく
        type: 'func',
        josi: [['から','を','の']],
        fn: function (cell, sys) {
            if (sys.tags.__workbook === null) {
                throw new Error(ERR_NO_WORKBOOK)
            }
            const objCell = sys.tags.__worksheet.getCell(cell)
            return objCell.value
        }
    },
    'エクセル取得': { // @セル(例えば「A1」)の値を取得して返す // @えくせるしゅとく
        type: 'func',
        josi: [['から','を','の']],
        fn: function (cell, sys) {
            return sys.__exec('エクセルセル取得', [cell, sys])
        }
    },
    'エクセル一括取得': { // @左上のセルC1(例えば「A1」)から右下のC2までの値を取得して二次元配列変数で返す // @えくせるいっかつしゅとく
        type: 'func',
        josi: [['から'],['までの', 'まで','の']],
        fn: function (c1, c2, sys) {
            if (sys.tags.__workbook === null) {
                throw new Error(ERR_NO_WORKBOOK)
            }
            const result = []
            const pos1 = Utils.addressToPos(c1)
            const pos2 = Utils.addressToPos(c2)
            // console.log(pos1)
            // console.log(pos2)
            for (let row = pos1.row; row <= pos2.row; row++) {
                const cells = []
                for (let col = pos1.col; col <= pos2.col; col++) {
                    const v = sys.tags.__worksheet.getRow(row).getCell(col).value
                    cells.push(v)
                }
                result.push(cells)
            }
            return result
        }
    },
    'エクセルシート列挙': { // @作業中のブックのシート一覧取得して返す // @えくせるしーとれっきょ
        type: 'func',
        josi: [],
        fn: function (sys) {
            if (sys.tags.__workbook === null) {
                throw new Error(ERR_NO_WORKBOOK)
            }
            const names = []
            sys.tags.__workbook.eachSheet(function(sheet, id){
                names.push(sheet.name)
            })
            return names
        }
    },
    'エクセルシート削除': { // @作業中のブックのシートNAMEを削除する // @えくせるしーとさくじょ
        type: 'func',
        josi: [['の','を']],
        fn: function (name, sys) {
            if (sys.tags.__workbook === null) {
                throw new Error(ERR_NO_WORKBOOK)
            }
            let sheet = sys.tags.__workbook.getWorksheet(name)
            if (!sheet) { throw new Error(`『EXCELシート削除』でシート『${name}』が見当たりません。`) }
            sys.tags.__workbook.removeWorksheet(sheet.id)
        },
        return_none: true
    },
    'エクセルセル幅設定': { // @作業中のシートcol列目の幅をWに設定する // @えくせるせるはばせってい
        type: 'func',
        josi: [['を'],['に','へ']],
        fn: function (col, w, sys) {
            if (sys.tags.__workbook === null) {
                throw new Error(ERR_NO_WORKBOOK)
            }
            sys.tags.__worksheet.getColumn(col).width = w
        },
        return_none: true
    },
    'エクセル背景色設定': { // @作業中シートのセルcells(例「A1」「A1:C3」)の背景色をcolorに設定 // @えくせるはいけいしょくせってい
        type: 'func',
        josi: [['を'],['に','へ']],
        fn: function (cells, color, sys) {
            if (sys.tags.__workbook === null) {
                throw new Error(ERR_NO_WORKBOOK)
            }
            const colorCode = Color.getColor(color)
            const range = Utils.addressToPosRange(cells)
            for (let row = range[0].row; row <= range[1].row; row++) {
                for (let col = range[0].col; col <= range[1].col; col++) {
                    const cell = sys.tags.__worksheet.getRow(row).getCell(col)
                    // cell.fill issue exceljs#791
                    cell.style = JSON.parse(JSON.stringify(cell.style))
                    // set fill
                    cell.style.fill = {
                        type: 'pattern',
                        pattern: 'solid',
                        fgColor: { argb: colorCode }
                    }
                }
                sys.tags.__worksheet.getRow(row).commit()
            }
        },
        return_none: true
    },
    'エクセル文字色設定': { // @作業中シートのセルcells(例「A1」「A1:C3」)の文字色をcolorに設定 // @えくせるもじいろせってい
        type: 'func',
        josi: [['を'],['に','へ']],
        fn: function (cells, color, sys) {
            if (sys.tags.__workbook === null) {
                throw new Error(ERR_NO_WORKBOOK)
            }
            const colorCode = Color.getColor(color)
            const range = Utils.addressToPosRange(cells)
            for (let row = range[0].row; row <= range[1].row; row++) {
                for (let col = range[0].col; col <= range[1].col; col++) {
                    const cell = sys.tags.__worksheet.getRow(row).getCell(col)
                    cell.style = JSON.parse(JSON.stringify(cell.style))
                    cell.font = {...cell.font, color: {argb: colorCode}}
                }
            }
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


