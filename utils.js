// Excel操作のための便利な命令をまとめたもの

const CAPITAL_A = 65;
function excelCoords(row, col) {
    var colStr = '';

    while(col > 0) {
        colStr = toChar((col - 1) % 26) + colStr;
        col = Math.floor((col - 1) / 26);
    }

    return colStr + row;
}

function toChar(n) {
    return String.fromCharCode(CAPITAL_A + n);
}

function cartesianCoords(excelCoords) {
    var row = parseInt(excelCoords.replace(/^[A-Z]+/, ''));
    var colChars = excelCoords.replace(/\d+$/, '').split('').reverse();
    var col = 0;
    var multiplier = 1;

    while(colChars.length) {
        col += toBase26Ish(colChars.shift()) * multiplier;
        multiplier *= 26;
    }

    return {row, col};
}

function toBase26Ish(c) {
    return c.charCodeAt(0) - CAPITAL_A + 1;
}

function posToAddress (row, col) {
    return excelCoords(row, col)
}

function addressToPos(addr) {
    return cartesianCoords(addr)
}

function addressToPosRange(addr) {
    addr = addr + ':'
    const addr2 = addr.split(':')
    if (addr2[1] == '') { addr2[1] = addr2[0] }
    const c1 = addressToPos(addr2[0])
    const c2 = addressToPos(addr2[1])
    return [c1, c2]
}

module.exports = {
    posToAddress, addressToPos, addressToPosRange
}
