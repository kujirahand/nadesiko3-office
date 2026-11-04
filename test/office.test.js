const { test } = require('node:test')
const assert = require('node:assert/strict')
const plugin = require('../nadesiko3-office.js')

function setup() {
    const variables = {}
    const sys = { tags: {}, __setSysVar(k, v) { variables[k] = v }, __getSysVar: k => variables[k] }
    sys.__exec = (name, args) => plugin[name].fn(...args)
    plugin['初期化'].fn(sys)
    return sys
}

test('初期化: システム変数と作業状態を初期化する', () => {
    const variables = {}
    const sys = { tags: {}, __setSysVar: (key, value) => { variables[key] = value } }
    plugin['初期化'].fn(sys)
    assert.equal(variables['OFFICEバージョン'], plugin.meta.value.pluginVersion)
    assert.equal(sys.tags.__workbook, null)
    assert.equal(sys.tags.__worksheet, null)
})

test('エクセル新規ブック: ブックと作業シートを生成する', () => {
    const sys = setup()
    const book = plugin['エクセル新規ブック'].fn(sys)
    assert.equal(sys.tags.__workbook, book)
    assert.equal(book.worksheets[0], sys.tags.__worksheet)
})

test('エクセル開: サンプルの最初のシートを選択する', async () => {
    const sys = setup()
    const book = await plugin['エクセル開'].fn(require('node:path').join(__dirname, '../sample/Book1.xlsx'), sys)
    assert.equal(sys.tags.__workbook, book)
    assert.equal(sys.tags.__worksheet, book.worksheets[0])
})

for (const [name, format] of [['エクセル保存', 'xlsx'], ['エクセルCSV保存', 'csv']]) {
    test(`${name}: 指定したファイル名で保存処理を呼ぶ`, async () => {
        const sys = setup()
        await assert.rejects(plugin[name].fn('output', sys), /ブック/)
        let filename
        sys.tags.__workbook = { [format]: { writeFile: file => { filename = file } } }
        await plugin[name].fn('output', sys)
        assert.equal(filename, 'output')
    })
}

test('エクセル新規シート: 指定名のシートを追加して選択する', () => {
    const sys = setup()
    plugin['エクセル新規ブック'].fn(sys)
    const sheet = plugin['エクセル新規シート'].fn('追加', sys)
    assert.equal(sheet.name, '追加')
    assert.equal(sys.tags.__worksheet, sheet)
})

test('エクセルシート取得: 指定名のシートを返す', () => {
    const sys = setup()
    plugin['エクセル新規ブック'].fn(sys)
    const sheet = plugin['エクセル新規シート'].fn('取得対象', sys)
    assert.equal(plugin['エクセルシート取得'].fn('取得対象', sys), sheet)
})

test('エクセルシート注目: 指定名のシートに切り替える', () => {
    const sys = setup()
    plugin['エクセル新規ブック'].fn(sys)
    const first = sys.tags.__worksheet
    plugin['エクセル新規シート'].fn('別', sys)
    assert.equal(plugin['エクセルシート注目'].fn(first.name, sys), first)
    assert.equal(sys.tags.__worksheet, first)
})

test('エクセルセル設定: 文字列と数式を設定する', () => {
    const sys = setup()
    plugin['エクセル新規ブック'].fn(sys)
    plugin['エクセルセル設定'].fn('A1', '値', sys)
    plugin['エクセルセル設定'].fn('A2', '=1+2', sys)
    assert.equal(sys.tags.__worksheet.getCell('A1').value, '値')
    assert.deepEqual(sys.tags.__worksheet.getCell('A2').value, { formula: '1+2' })
})

test('エクセル設定: セルに値を設定する', () => {
    const sys = setup()
    plugin['エクセル新規ブック'].fn(sys)
    plugin['エクセル設定'].fn('B1', '別名', sys)
    assert.equal(sys.tags.__worksheet.getCell('B1').value, '別名')
})

test('エクセル一括設定: 指定位置から配列を設定する', () => {
    const sys = setup()
    plugin['エクセル新規ブック'].fn(sys)
    plugin['エクセル一括設定'].fn('B2:C3', [[1, 2], [3, 4]], sys)
    assert.equal(sys.tags.__worksheet.getCell('B2').value, 1)
    assert.equal(sys.tags.__worksheet.getCell('C3').value, 4)
})

test('エクセルセル取得: 指定セルの値を返す', () => {
    const sys = setup()
    plugin['エクセル新規ブック'].fn(sys)
    sys.tags.__worksheet.getCell('A1').value = 42
    assert.equal(plugin['エクセルセル取得'].fn('A1', sys), 42)
})

test('エクセル取得: 指定セルの値を返す', () => {
    const sys = setup()
    plugin['エクセル新規ブック'].fn(sys)
    sys.tags.__worksheet.getCell('B2').value = '値'
    assert.equal(plugin['エクセル取得'].fn('B2', sys), '値')
})

test('エクセル一括取得: 指定範囲を二次元配列で返す', () => {
    const sys = setup()
    plugin['エクセル新規ブック'].fn(sys)
    sys.tags.__worksheet.getCell('C2').value = 7
    assert.deepEqual(plugin['エクセル一括取得'].fn('B1', 'C2', sys), [[null, null], [null, 7]])
})

test('エクセルシート列挙: シート名を順に返す', () => {
    const sys = setup()
    plugin['エクセル新規ブック'].fn(sys)
    const first = sys.tags.__worksheet.name
    plugin['エクセル新規シート'].fn('次', sys)
    assert.deepEqual(plugin['エクセルシート列挙'].fn(sys), [first, '次'])
})

test('エクセルシート削除: 指定シートを削除し、存在しない名前はエラーになる', () => {
    const sys = setup()
    plugin['エクセル新規ブック'].fn(sys)
    plugin['エクセル新規シート'].fn('削除対象', sys)
    plugin['エクセルシート削除'].fn('削除対象', sys)
    assert.equal(sys.tags.__workbook.getWorksheet('削除対象'), undefined)
    assert.throws(() => plugin['エクセルシート削除'].fn('削除対象', sys), /見当たりません/)
})

test('エクセルセル幅設定: 指定列の幅を変更する', () => {
    const sys = setup()
    plugin['エクセル新規ブック'].fn(sys)
    plugin['エクセルセル幅設定'].fn('B', 25, sys)
    assert.equal(sys.tags.__worksheet.getColumn('B').width, 25)
})

test('エクセル背景色設定: 指定範囲に背景色を設定する', () => {
    const sys = setup()
    plugin['エクセル新規ブック'].fn(sys)
    plugin['エクセル背景色設定'].fn('A1:B2', '赤', sys)
    assert.equal(sys.tags.__worksheet.getCell('A1').fill.fgColor.argb, 'FFff0000')
    assert.equal(sys.tags.__worksheet.getCell('B2').fill.fgColor.argb, 'FFff0000')
})

test('エクセル文字色設定: 指定範囲に文字色を設定する', () => {
    const sys = setup()
    plugin['エクセル新規ブック'].fn(sys)
    plugin['エクセル文字色設定'].fn('A1:B2', '赤', sys)
    assert.equal(sys.tags.__worksheet.getCell('A1').font.color.argb, 'FFff0000')
    assert.equal(sys.tags.__worksheet.getCell('B2').font.color.argb, 'FFff0000')
})

test('エクセル行数取得: 空のシートでは0を返す', () => {
    const sys = setup()
    plugin['エクセル新規ブック'].fn(sys)
    assert.equal(plugin['エクセル行数取得'].fn(sys), 0)
})

test('エクセル行数取得: 間に空行があっても最終行番号を返す', () => {
    const sys = setup()
    plugin['エクセル新規ブック'].fn(sys)
    plugin['エクセルセル設定'].fn('B5', '末尾', sys)
    assert.equal(plugin['エクセル行数取得'].fn(sys), 5)
})

test('エクセルシート全取得: 空のシートでは空配列を返す', () => {
    const sys = setup()
    plugin['エクセル新規ブック'].fn(sys)
    assert.deepEqual(plugin['エクセルシート全取得'].fn(sys), [])
})

test('エクセルシート全取得: 空行と空セルを含む二次元配列を返す', () => {
    const sys = setup()
    plugin['エクセル新規ブック'].fn(sys)
    plugin['エクセルセル設定'].fn('B3', '値', sys)
    assert.deepEqual(plugin['エクセルシート全取得'].fn(sys), [
        [null, null], [null, null], [null, '値']
    ])
})

test('行数とシート全取得はブックとシートを必要とする', () => {
    const sys = setup()
    for (const name of ['エクセル行数取得', 'エクセルシート全取得']) {
        assert.throws(() => plugin[name].fn(sys), /ブック/)
        sys.tags.__workbook = {}
        assert.throws(() => plugin[name].fn(sys), /シート/)
        sys.tags.__workbook = null
    }
})

test('空シート、空行、空セルと各種セル値を取得する', () => {
    const sys = setup()
    plugin['エクセル新規ブック'].fn(sys)
    const sheet = sys.tags.__worksheet
    const count = () => plugin['エクセル行数取得'].fn(sys)
    const all = () => plugin['エクセルシート全取得'].fn(sys)
    assert.equal(count(), 0)
    assert.deepEqual(all(), [])
    sheet.getCell('B1').value = 0
    sheet.getCell('C3').value = { formula: '1+1', result: 2 }
    assert.equal(count(), 3)
    assert.deepEqual(all(), [
        [null, 0, null],
        [null, null, null],
        [null, null, { formula: '1+1', result: 2 }]
    ])
    assert.equal(sheet.findRow(2), undefined)
    plugin['エクセル新規シート'].fn('別シート', sys)
    assert.equal(count(), 0)
    assert.deepEqual(all(), [])
})

test('Excelファイルの読込と保存後も値を取得できる', async () => {
    const sys = setup()
    await plugin['エクセル開'].fn(require('node:path').join(__dirname, '../sample/Book1.xlsx'), sys)
    const before = plugin['エクセルシート全取得'].fn(sys)
    assert.ok(before.length > 0)
    const buffer = await sys.tags.__workbook.xlsx.writeBuffer()
    const workbook = new (require('exceljs').Workbook)()
    await workbook.xlsx.load(buffer)
    sys.tags.__workbook = workbook
    sys.tags.__worksheet = workbook.worksheets[0]
    assert.deepEqual(plugin['エクセルシート全取得'].fn(sys), before)
})

test('シート切替後は選択中のシートの内容を取得する', () => {
    const sys = setup()
    plugin['エクセル新規ブック'].fn(sys)
    const firstName = sys.tags.__worksheet.name
    plugin['エクセルセル設定'].fn('A2', '最初', sys)
    plugin['エクセル新規シート'].fn('二枚目', sys)
    plugin['エクセルセル設定'].fn('B1', '二枚目の値', sys)
    assert.equal(plugin['エクセル行数取得'].fn(sys), 1)
    assert.deepEqual(plugin['エクセルシート全取得'].fn(sys), [[null, '二枚目の値']])
    plugin['エクセルシート注目'].fn(firstName, sys)
    assert.equal(plugin['エクセル行数取得'].fn(sys), 2)
    assert.deepEqual(plugin['エクセルシート全取得'].fn(sys), [[null], ['最初']])
})

test('一括設定した値と数式は個別取得と同じ形式で返る', () => {
    const sys = setup()
    plugin['エクセル新規ブック'].fn(sys)
    plugin['エクセル一括設定'].fn('B2', [[42, false], ['文字', null]], sys)
    plugin['エクセルセル設定'].fn('A1', '=1+1', sys)
    const result = plugin['エクセルシート全取得'].fn(sys)
    assert.deepEqual(result, [
        [{ formula: '1+1' }, null, null],
        [null, 42, false],
        [null, '文字', null]
    ])
    assert.deepEqual(result[0][0], plugin['エクセルセル取得'].fn('A1', sys))
    assert.deepEqual(result.slice(1).map(row => row.slice(1)),
        plugin['エクセル一括取得'].fn('B2', 'C3', sys))
})

test('書式だけの行も行数に含め、取得はシートを変更しない', () => {
    const sys = setup()
    plugin['エクセル新規ブック'].fn(sys)
    const sheet = sys.tags.__worksheet
    sheet.getCell('C4').font = { bold: true }
    assert.equal(plugin['エクセル行数取得'].fn(sys), 4)
    const expected = [[], [], [], []]
    assert.deepEqual(plugin['エクセルシート全取得'].fn(sys), expected)
    assert.deepEqual(plugin['エクセルシート全取得'].fn(sys), expected)
    assert.equal(sheet.findRow(1), undefined)
    assert.equal(sheet.getCell('C4').font.bold, true)
})

test('エクセルシート全取得: セル数が上限を超えるとエラーになる', () => {
    const sys = setup()
    assert.equal(sys.__getSysVar('エクセル全取得最大セル数'), 1000000)
    plugin['エクセル新規ブック'].fn(sys)
    plugin['エクセルセル設定'].fn('C2', '値', sys)
    sys.__setSysVar('エクセル全取得最大セル数', 5)
    assert.throws(() => plugin['エクセルシート全取得'].fn(sys), /大きすぎます/)
    sys.__setSysVar('エクセル全取得最大セル数', 6)
    assert.equal(plugin['エクセルシート全取得'].fn(sys).length, 2)
})
