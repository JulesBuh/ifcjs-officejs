/*jsoffice script functions devised by Julian Buhagiar, incorporating RFC 4122, 9562 and ISO 9348, 16739 guid methods.*/

/** Returns 32byte guid with dashes 01234567-89ab-<Method>def-a123-456789abcdef- */
/** @CustomFunction
 * @param {string} guid - input a guid e.g 0123456789abcdef0123456789abcdef
 *  * @param {string} method - 0 to f
 * @description adds dashes
 * @returns 32 guid with dashes 01234567-89ab-<Method>def-a123-456789abcdef
 * @volatile
 */
function AddDashes(guid, method = guid.substr(12, 1)) {
  /*ensure no dashes are already present*/
  if (/[a-fA-F0-9]{8}-[a-fA-F0-9]{4}-[a-fA-F0-9]{4}-[a-fA-F0-9]{4}-[a-fA-F0-9]{12}/.test(guid)) {
    guid = RemoveDashes(guid, method);
  }
  /*then adds dashes*/
  if (/[a-fA-F0-9]{32}/.test(guid)) {
    var temp = "";
    //  the regex pattern matches and will add dashes to the 32 length string
    temp = temp.concat(
      guid.substr(0, 8),
      "-",
      guid.substr(8, 4),
      "-",
      method,
      guid.substr(13, 3),
      "-",
      guid.substr(16, 4),
      "-",
      guid.substr(20, 12)
    );
    //console.log('successfully added dashes in the format XXXXXXXX-XXXX-' + method + 'XXX-XXXX-XXXXXXXXXXXX'.');
    return temp;
  } else {
    temp = "00000000-0000-" + method + "000-0000-00000000000";
    //console.log('something not right with "' + string + '".');
    return;
  }
}
function RemoveDashes(guid, method = guid.substr(12, 1)) {
  var temp = "";
  if (/[a-fA-F0-9]{8}-[a-fA-F0-9]{4}-[a-fA-F0-9]{4}-[a-fA-F0-9]{4}-[a-fA-F0-9]{12}/.test(guid)) {
    //console.log('true')
    temp = guid.replace(/-/g, "");
  }
  //console.log('successfully removed dashes from input "XXXXXXXXXXXX" + method + "XXXXXXXXXXXXXXXXXXX".',temp);
  return temp;
}
function RemoveBrackets(guid) {
  var temp = "";
  if (/\{[a-fA-F0-9]{8}-[a-fA-F0-9]{4}-[a-fA-F0-9]{4}-[a-fA-F0-9]{4}-[a-fA-F0-9]{12}\}/.test(guid)) {
    temp = guid.replace(/\{/g, "").replace(/\}/g, "");
  }
}

/** Ifc guid Compression Function - Converts Ifc guid to hex guid- */
/** @CustomFunction
 * @param {string} ifcGUID - input a compressed guid e.g 0123789ABCDXYZabcxyz_$
 * @description Converts hex guid to ifc guid.
 * @returns 32 character hex guid.á
 * @volatile
 */
function IfcToGuid(ifcGUID) {
  var str = ifcGUID;
  var ifcGuidPatt = new RegExp("^[0-3][a-zA-Z0-9_$]{21}$", "gi"); //Checked this regex exhaustively in Excel by applying the formula to different permutations of characters and lengths. Valid ifcguids must begin with 0,1,2 or 3.
  var res = str.match(ifcGuidPatt);
  function convertBinaryToHexadecimal(binaryString) {
    var output = "";

    // For every 4 bits in the binary string
    for (var i = 0; i < binaryString.length; i += 4) {
      // Grab a chunk of 4 bits
      var bytes = binaryString.substr(i, 4);

      // Convert to decimal then hexadecimal
      var decimal = parseInt(bytes, 2);
      var hex = decimal.toString(16);

      // Uppercase all the letters and append to output
      output += hex.toUpperCase();
    }

    return output;
  }
  if (res) {
    //perform conversion if valid ifcguid
    var ifcGuidChar = new RegExp("[a-zA-Z0-9_$]", "gi");
    var res = str.match(ifcGuidChar);
    //return res.toString();
    for (var x = 0; x < str.length; x++) {
      var c = str.charAt(x);
      //code here to do the binary translation
      //64bit to bin key value pairs
      var arr = {
        "0": "000000",
        "1": "000001",
        "2": "000010",
        "3": "000011",
        "4": "000100",
        "5": "000101",
        "6": "000110",
        "7": "000111",
        "8": "001000",
        "9": "001001",
        A: "001010",
        B: "001011",
        C: "001100",
        D: "001101",
        E: "001110",
        F: "001111",
        G: "010000",
        H: "010001",
        I: "010010",
        J: "010011",
        K: "010100",
        L: "010101",
        M: "010110",
        N: "010111",
        O: "011000",
        P: "011001",
        Q: "011010",
        R: "011011",
        S: "011100",
        T: "011101",
        U: "011110",
        V: "011111",
        W: "100000",
        X: "100001",
        Y: "100010",
        Z: "100011",
        a: "100100",
        b: "100101",
        c: "100110",
        d: "100111",
        e: "101000",
        f: "101001",
        g: "101010",
        h: "101011",
        i: "101100",
        j: "101101",
        k: "101110",
        l: "101111",
        m: "110000",
        n: "110001",
        o: "110010",
        p: "110011",
        q: "110100",
        r: "110101",
        s: "110110",
        t: "110111",
        u: "111000",
        v: "111001",
        w: "111010",
        x: "111011",
        y: "111100",
        z: "111101",
        _: "111110",
        "[$]": "111111"
      };
      //translate function 64bit to bin string
      var bin_str = str;
      for (var key in arr) {
        if (!arr.hasOwnProperty(key)) {
          continue;
        }
        bin_str = bin_str.replace(new RegExp("" + key, "g"), arr[key]);
      }
      //unpad binary string (cull first four bits '0000')
      var bin_str = bin_str.substring(4, 132);
      //break down binary string into 4-bit chunks
      var hex_guid = "";
      for (var i = 0; i < 128; i += 4) {
        hex_guid += convertBinaryToHexadecimal(bin_str.substr(i, 4));
      }
      hex_guid = hex_guid.toLowerCase();
      hex_guid =
        hex_guid.substr(0, 8) +
        "-" +
        hex_guid.substr(8, 4) +
        "-" +
        hex_guid.substr(12, 4) +
        "-" +
        hex_guid.substr(16, 4) +
        "-" +
        hex_guid.substr(20, 12);
    }
    return hex_guid.toString();
  } else {
    //don't perform conversion
    return "#N/A";
  }
}
/** Ifc guid Compression Function - Converts hex guid to ifc guid- */
/** @CustomFunction
 * @param {string} hexGUID - input a 128-bit guid e.g 01234567-89ab-0cde-f012-3456789abcde
 * @description Converts ifc guid to hex guid.
 * @returns 22 character 64base guid.
 * @volatile
 */
function GuidToIfc(hexGUID) {
  var str = hexGUID;
  var ifcGuidPatt = new RegExp("^[a-fA-F0-9]{32}$|^[a-fA-F0-9-]{36}$|^{[a-fA-F0-9-]{36}}$", "gi");
  //Checked this regex exhaustively in Excel by applying the formula to different permutations of characters and lengths
  var res = str.match(ifcGuidPatt);
  function convertBinaryToHexadecimal(binaryString) {
    var output = "";

    // For every 6 bits in the binary string
    for (var i = 0; i < binaryString.length; i += 6) {
      // Grab a chunk of 6 bits
      var bytes = binaryString.substr(i, 6);

      // Convert to decimal then hexadecimal
      var decimal = parseInt(bytes, 2) + "|";

      var arr2 = {
        "000000": "0",
        "000001": "1",
        "000010": "2",
        "000011": "3",
        "000100": "4",
        "000101": "5",
        "000110": "6",
        "000111": "7",
        "001000": "8",
        "001001": "9",
        "001010": "A",
        "001011": "B",
        "001100": "C",
        "001101": "D",
        "001110": "E",
        "001111": "F",
        "010000": "G",
        "010001": "H",
        "010010": "I",
        "010011": "J",
        "010100": "K",
        "010101": "L",
        "010110": "M",
        "010111": "N",
        "011000": "O",
        "011001": "P",
        "011010": "Q",
        "011011": "R",
        "011100": "S",
        "011101": "T",
        "011110": "U",
        "011111": "V",
        "100000": "W",
        "100001": "X",
        "100010": "Y",
        "100011": "Z",
        "100100": "a",
        "100101": "b",
        "100110": "c",
        "100111": "d",
        "101000": "e",
        "101001": "f",
        "101010": "g",
        "101011": "h",
        "101100": "i",
        "101101": "j",
        "101110": "k",
        "101111": "l",
        "110000": "m",
        "110001": "n",
        "110010": "o",
        "110011": "p",
        "110100": "q",
        "110101": "r",
        "110110": "s",
        "110111": "t",
        "111000": "u",
        "111001": "v",
        "111010": "w",
        "111011": "x",
        "111100": "y",
        "111101": "z",
        "111110": "_",
        "111111": "$"
      };
      // Uppercase all the letters and append to output
      output += arr2[binaryString.substr(i, 6)];
    }

    return output;
  }
  if (res) {
    //perform conversion if valid ifcguid
    var res = str.replace(/\-/g, "");
    //
    var arr = {
      "0": "0000",
      "1": "0001",
      "2": "0010",
      "3": "0011",
      "4": "0100",
      "5": "0101",
      "6": "0110",
      "7": "0111",
      "8": "1000",
      "9": "1001",
      A: "1010",
      B: "1011",
      C: "1100",
      D: "1101",
      E: "1110",
      F: "1111",
      a: "1010",
      b: "1011",
      c: "1100",
      d: "1101",
      e: "1110",
      f: "1111"
    };
    var bin_str = res;
    for (var key in arr) {
      if (!arr.hasOwnProperty(key)) {
        continue;
      }
      bin_str = bin_str.replace(new RegExp("" + key, "g"), arr[key]);
    }
    //pad
    bin_str = "0000" + bin_str;
    var ifc_guid = "";

    for (var i = 0; i < 128; i += 6) {
      ifc_guid += convertBinaryToHexadecimal(bin_str.substr(i, 6));
    }

    return ifc_guid.toString();
  } else {
    //don't perform conversion
    return "#N/A";
  }
}
