// define colors
const Colors = {
    "黒":"000000", "白":"ffffff", "赤": "ff0000",
    "青": "0000ff", "黄": "ffff00", "紫": "ff00ff",
    "緑": "00ff00",
    "black":"000000","aliceblue":"f0f8ff","darkcyan":"008b8b","lightyellow":"ffffe0","coral":"ff7f50","dimgray":"696969","lavender":"e6e6fa","teal":"008080","lightgoldenrodyellow":"fafad2","tomato":"ff6347","gray":"808080","lightsteelblue":"b0c4de","darkslategray":"2f4f4f","lemonchiffon":"fffacd","orangered":"ff4500","darkgray":"a9a9a9","lightslategray":"778899","darkgreen":"006400","wheat":"f5deb3","red":"ff0000","silver":"c0c0c0","slategray":"708090","green":"008000","burlywood":"deb887","crimson":"dc143c","lightgray":"d3d3d3","steelblue":"4682b4","forestgreen":"228b22","tan":"d2b48c","mediumvioletred":"c71585","gainsboro":"dcdcdc","royalblue":"4169e1","seagreen":"2e8b57","khaki":"f0e68c","deeppink":"ff1493","whitesmoke":"f5f5f5","midnightblue":"191970","mediumseagreen":"3cb371","yellow":"ffff00","hotpink":"ff69b4","white":"ffffff","navy":"000080","mediumaquamarine":"66cdaa","gold":"ffd700","palevioletred":"db7093","snow":"fffafa","darkblue":"00008b","darkseagreen":"8fbc8f","orange":"ffa500","pink":"ffc0cb","ghostwhite":"f8f8ff","mediumblue":"0000cd","aquamarine":"7fffd4","sandybrown":"f4a460","lightpink":"ffb6c1","floralwhite":"fffaf0","blue":"0000ff","palegreen":"98fb98","darkorange":"ff8c00","thistle":"d8bfd8","linen":"faf0e6","dodgerblue":"1e90ff","lightgreen":"90ee90","goldenrod":"daa520","magenta":"ff00ff","antiquewhite":"faebd7","cornflowerblue":"6495ed","springgreen":"00ff7f","peru":"cd853f","fuchsia":"ff00ff","papayawhip":"ffefd5","deepskyblue":"00bfff","mediumspringgreen":"00fa9a","darkgoldenrod":"b8860b","violet":"ee82ee","blanchedalmond":"ffebcd","lightskyblue":"87cefa","lawngreen":"7cfc00","chocolate":"d2691e","plum":"dda0dd","bisque":"ffe4c4","skyblue":"87ceeb","chartreuse":"7fff00","sienna":"a0522d","orchid":"da70d6","moccasin":"ffe4b5","lightblue":"add8e6","greenyellow":"adff2f","saddlebrown":"8b4513","mediumorchid":"ba55d3","navajowhite":"ffdead","powderblue":"b0e0e6","lime":"00ff00","maroon":"800000","darkorchid":"9932cc","peachpuff":"ffdab9","paleturquoise":"afeeee","limegreen":"32cd32","darkred":"8b0000","darkviolet":"9400d3","mistyrose":"ffe4e1","lightcyan":"e0ffff","yellowgreen":"9acd32","brown":"a52a2a","darkmagenta":"8b008b","lavenderblush":"fff0f5","cyan":"00ffff","darkolivegreen":"556b2f","firebrick":"b22222","purple":"800080","seashell":"fff5ee","aqua":"00ffff","olivedrab":"6b8e23","indianred":"cd5c5c","indigo":"4b0082","oldlace":"fdf5e6","turquoise":"40e0d0","olive":"808000","rosybrown":"bc8f8f","darkslateblue":"483d8b","ivory":"fffff0","mediumturquoise":"48d1cc","darkkhaki":"bdb76b","darksalmon":"e9967a","blueviolet":"8a2be2","honeydew":"f0fff0","darkturquoise":"00ced1","palegoldenrod":"eee8aa","lightcoral":"f08080","mediumpurple":"9370db","mintcream":"f5fffa","lightseagreen":"20b2aa","cornsilk":"fff8dc","salmon":"fa8072","slateblue":"6a5acd","azure":"f0ffff","cadetblue":"5f9ea0","beige":"f5f5dc","lightsalmon":"ffa07a","mediumslateblue":"7b68ee"
}
// get color code
function getColor(color) {
    color = '' + color
    color = color.toLowerCase().replace(/\s/g, '')
    if (Colors[color]) { 
        color = 'FF' + Colors[color] // argb
    }
    color = color.replace(/^\#/, '')
    return color
}

module.exports = { getColor }
