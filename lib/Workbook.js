"use strict";

var fs = require('fs');
var JSZip = require('jszip');
var utils = require('./utils');
var Sheet = require('./Sheet');
var path = require("path");
var DOMParser = require('xmldom').DOMParser;
var parser = new DOMParser();
var xpath = require("./xpath");

/**
 * Initializes a new Workbook.
 * @param {Buffer} [data] - File buffer of the Excel workbook (optional).
 * @constructor
 */
var Workbook = function (data) {
    /*
    This base64 string can be generated with the following code:

    var fs = require('fs');
    var blankDataString = fs
        .readFileSync('./lib/blank.xlsx')
        .toString('base64')
        ;
    */
    data = data || Buffer.from(
        'UEsDBBQAAAAIAAAAIQC1VTAj7AAAAEwCAAALAAAAX3JlbHMvLnJlbHONks1OwzAMgO9IvEPk++puSAihpbsgpN0QKg9gEvdHbeMoCdC9PeGAoNIYPcaxP3+2vD/M06jeOcRenIZtUYJiZ8T2rtXwUj9u7kDFRM7SKI41nDjCobq+2j/zSCkXxa73UWWKixq6lPw9YjQdTxQL8ezyTyNhopSfoUVPZqCWcVeWtxh+M6BaMNXRaghHewOqPnlew5am6Q0/iHmb2KUzLZDnxM6y3fiQ60Pq8zSqptBy0mDFPOVwRPK+yGjA80a79UZ/T4sTJ7KUCI0EvuzzlXFJaLte6P8VLTN+bOYRPyQMryLDtwsubqD6BFBLAwQUAAAACAAAACEA3kEW2XsBAAARAwAAEAAAAGRvY1Byb3BzL2FwcC54bWydkkFP4zAQhe9I/IfId+oElhWqHCNUQBwWbaUWOBtn0lg4tuUZopZfj5OqIV32xO3NzNPLlxmL621rsw4iGu9KVsxyloHTvjJuU7Kn9f3ZFcuQlKuU9Q5KtgNk1/L0RCyjDxDJAGYpwmHJGqIw5xx1A63CWRq7NKl9bBWlMm64r2uj4dbr9xYc8fM8/81hS+AqqM7CGMj2ifOOfhpaed3z4fN6F1KeFDchWKMVpb+Uj0ZHj76m7G6rwQo+HYoUtAL9Hg3tZC74tBQrrSwsUrCslUUQ/KshHkD1S1sqE1GKjuYdaPIxQ/OR1nbOsleF0OOUrFPRKEdsb9sXg7YBKcoXH9+wASAUfGwOcuqdavNLFoMhiWMjH0GSPkZcG7KAf+ulivQf4mJKPDCwCeOq5yu+8R2+9E/2wrdBubRAPqo/xr3hU1j7W0VwWOdxU6waFaFKFxjXPTbEQ+KKtvcvGuU2UB083wf98Z/3L1wWl7P8Is+Hmx96gn+9ZfkJUEsDBBQAAAAIAOehdkc+qGWw1QAAAG0BAAARAAAAZG9jUHJvcHMvY29yZS54bWxtkE1Lw0AQhu9C/0PYezKJBZGQpDdPCkIVvA67Y7qY/WBnNO2/7zZoFOxxeJ95mHm73dFNxRcltsH3qqlqVZDXwVg/9ur15aG8VwULeoNT8NSrE7HaDZubTsdWh0TPKURKYomLbPLc6tirg0hsAVgfyCFXmfA5fA/JoeQxjRBRf+BIcFvXd+BI0KAgXIRlXI3qW2n0qoyfaVoERgNN5MgLQ1M18MsKJcdXF5bkD+msnCJdRX/ClT6yXcF5nqt5u6D5/gbenh73y6ul9ZeuNKmhg38FDWdQSwMEFAAAAAAA2aF2RwAAAAAAAAAAAAAAAAkAAAB4bC9fcmVscy9QSwMEFAAAAAgAAAAhAI2H2nDaAAAALQIAABoAAAB4bC9fcmVscy93b3JrYm9vay54bWwucmVsc62R3YrCMBCF7xf2HcLcb9NWWGQx9UYWeiv1AUI6/cE2CZlZtW9vXMEfEPHCq+FMmO+cySyWh3EQOwzUO6sgS1IQaI2re9sq2FS/X3MQxNrWenAWFUxIsCw+PxZrHDTHIep6TyJSLCnomP2PlGQ6HDUlzqONL40Lo+YoQyu9NlvdoszT9FuGWwYUd0xR1gpCWc9AVJPHV9iuaXqDK2f+RrT8wEIST0NcQFQ6tMgKzjqJHJCP7fN32nOcxav7vzw3s2cZsndm2LuwpQ6RrzkurfhBp3IJI++OXBwBUEsDBBQAAAAIAAAAIQDeI/LTbgIAALEFAAANAAAAeGwvc3R5bGVzLnhtbKWUXWvbMBSG7wf7D0L3rmw3zpJguyxNDYVuDJrBbhVbTkT1YSSlSzb233tkO7FDxzbWK53z6ug5rz7s9OYgBXpmxnKtMhxdhRgxVeqKq22Gv66LYIaRdVRVVGjFMnxkFt/k79+l1h0Fe9wx5hAglM3wzrlmQYgtd0xSe6UbpmCm1kZSB6nZEtsYRivrF0lB4jCcEkm5wh1hIct/gUhqnvZNUGrZUMc3XHB3bFkYyXJxv1Xa0I0Aq4doQssTu01e4SUvjba6dleAI7quecleu5yTOQFSntZaOYtKvVcOzgrQHrp4Uvq7KvyUF7uqPLU/0DMVoESY5GmphTbIQVfmi0BRVLKu4pYKvjHcizWVXBw7OfZCa7Svkxy25kXSdWgHC4u4EGdXMe6EPIXTccyoAhLUx+tjA+0VXGSHaev+Ur019BjFyWhBO0DfjTYVPJzhPE5SngpWO1hg+HbnR6cb4iedg1PO04rTrVZUeORpRR8AtmRCPPrH9a2+YB9qpPaykO6+yjA8U7/7UwiG+rDDdInnj2kd+81YdKgv+Wd02+iCflaRv+8Mf/YPWQwItNlz4bj6jWFgVofBazvr/Mu+7AKMitV0L9z6PJnhIf7EKr6X8bnqC3/Wrq8a4gd/U9HU92AH92BdO6K94Rn+ebf8MF/dFXEwC5ezYHLNkmCeLFdBMrldrlbFPIzD21+jD+0Nn1n7O4BLiSYLK6DK9JvtzT8OWoZHSWe/PT+wPfY+j6fhxyQKg+I6jILJlM6C2fQ6CYokilfTyfIuKZKR9+T/vEchiaLBfLJwXDLBFbu0vx6rcEmQ/mET5HQTZPjX5i9QSwMEFAAAAAAA2aF2RwAAAAAAAAAAAAAAAAkAAAB4bC90aGVtZS9QSwMEFAAAAAgAAAAhAIuCblj1BQAAjhoAABMAAAB4bC90aGVtZS90aGVtZTEueG1s7VlPjxs1FL8j8R2suafzfyZZNVslk6SF7rZVd1vUozNxMm4842js7G5UVULtEQkJURAXJG4cEFCplbiUT7NQBEXqV8DjyR9P4tCFbqWCmkjJ+Pn3nn9+7/nZM3Px0klKwBHKGaZZ07AvWAZAWUwHOBs1jVuHvVrdAIzDbAAJzVDTmCFmXNp9/72LcIcnKEVA6GdsBzaNhPPJjmmyWIghu0AnKBN9Q5qnkItmPjIHOTwWdlNiOpYVmCnEmQEymAqz14dDHCNwWJg0dhfGu0T8ZJwVgpjkB7EcUdWQ2MHYLv7YjEUkB0eQNA0xzoAeH6ITbgACGRcdTcOSH8PcvWgulQjfoqvo9eRnrjdXGIwdqZeP+ktFz/O9oLW075T2N3HdsBt0g6U9CYBxLGZqb2D9dqPd8edYBVReamx3wo5rV/CKfXcD3/KLbwXvrvDeBr7Xi1Y+VEDlpa/xSehEXgXvr/DBBj60Wh0vrOAlKCE4G2+gLT9wo8Vsl5AhJVe08Ibv9UJnDl+hTCW7Sv2Mb8u1FN6leU8AZHAhxxngswkawljgIkhwP8dgD48SkXgTmFEmxJZj9SxX/BZfT15Jj8AdBBXtUhSzDVHBB7A4xxPeND4UVg0F8vLZ9y+fPQEvnz0+ffD09MFPpw8fnj74UaN4BWYjVfHFt5/9+fXH4I8n37x49IUez1T8rz988svPn+uBXAU+//Lxb08fP//q09+/e6SBt3LYV+GHOEUMXEPH4CZNxdw0A6B+/s80DhOIKxowEUgNsMuTCvDaDBIdro2qzrudiyKhA16e3q1wPUjyKcca4NUkrQD3KSVtmmunc7UYS53ONBvpB8+nKu4mhEe6saO10HanE5HtWGcySlCF5g0iog1HKEMcFH10jJBG7Q7GFb/u4zinjA45uINBG2KtSw5xn+uVruBUxGWmIyhCXfHN/m3QpkRnvoOOqkixICDRmUSk4sbLcMphqmUMU6Ii9yBPdCQPZnlccTjjItIjRCjoDhBjOp3r+axC96ooLvqw75NZWkXmHI91yD1IqYrs0HGUwHSi5YyzRMV+wMYiRSG4QbmWBK2ukKIt4gCzreG+jVEl3K9e1rdEXdUnSNEzzXVLAtHqepyRIUTSuLlWzVOcvbK0rxV1/11R1xf1Vo61S2u9lG/D/QcLeAdOsxtIrBkN9F39fle///f1e9taPv+qvSrUZqmonN3TrUf3ISbkgM8I2mOyxDMxvUFPCGVDKi3vFCaJuJwPV8GNciivQU75R5gnBwmciGFsOcKIzU2PGJhQJjYJY6vtooNM0306KKW2vbg5FQqQr+Rik1nIxZbES2kQru7CluZla8RUAr40enYSymBVEq6GROiejYRtnReLhoZF3f47FqYSFbH+ACyea/heyUjkGyRoUMSp1F9E99wjvc2Z1Wk7muk1vLM5+QyRrpBQ0q1KQknDBA7QuvicY91YhbRCz9HSCOtvItbmZm0gWbUFjsWac31hJoaTpjEUx0NxmU6EPVbUTUhGWdOI+dzR/6ayTHLGO5AlJUx2lfNPMUc5IDgVua6GgWQrbrYTWm8vuYb19nnOXA8yGg5RzLdIVk3RVxrR9r4muGjQqSB9kAyOQZ9M85tQOMoP7cKBA8z40psDnCvJvfLiWrmaL8XKQ7PVEoVkksD5jqIW8xIur5d0lHlIpuuzqrbnk+mPeuex675aqehQiuaWDSTcWsXe3CavsHL1rHxtrWvUl1L9LvH6G4JCra6n5uqpWVuoneOBQBku2OK35R5x3rvBetaayrlStjbeTtD+XZH5HXFcnRLOJFV0Iu4RosVz5bISSOmiupxwMM1x07hn+S0vcvyoZtX9bs1zPatW91tureX7rt31bavTdu4Lp/Aktf1y7J64nyGz+csXKd94AZMujtkXYpqaVJ6DTaksX8DYzvYXMAALz9wLnF7DbbSDWsNt9Wpep12vNaKgXesEUdjpdSK/3ujdN8CRBHstN/KCbr0W2FFU8wKroF9v1ELPcVpe2Kp3vdb9ua/FzBf/C/dKXrt/AVBLAwQUAAAACAAAACEAfDzuwy4CAACbBAAADwAAAHhsL3dvcmtib29rLnhtbK2UTY+bMBCG75X6H5DvhI9AN0Ehq81H1UjVarXN7l5yccwQ3Bib2qZJVPW/d4CSps1lK+0Fj8344Z13bCa3x1I430EbrmRKgoFPHJBMZVzuUvK0/uiOiGMslRkVSkJKTmDI7fT9u8lB6f1Wqb2DAGlSUlhbJZ5nWAElNQNVgcQ3udIltTjVO89UGmhmCgBbCi/0/Q9eSbkkHSHRr2GoPOcMForVJUjbQTQIalG+KXhlelrJXoMrqd7XlctUWSFiywW3pxZKnJIlq51Umm4Fln0M4p6M4RW65Ewro3I7QNRvkVf1Br4XBF3J00nOBTx3tju0qu5p2XxFEEdQY5cZt5ClBGUIdYC/FnRdzWoucBJEUegTb3puxYN2MshpLewaZfV4TIyHYRg2mVjUnbCgJbUwV9Kih2/kV8ueFwoLdx7hW801mM626QSflCV0ax6oLZxai5TMk82TQX2b7KsqpFFyM1cZbI7CHN1KVTV2FDaCbzcXrtNrif/hO2WNAd5ZZRf/68Z00hj5zOFg/vjaTJ3jC5eZOqRkPMQ7cupnGB/a8IVntkhJOPRvzmufgO8Kiw3wh12nvAt6K7AfHdkegC9NHOCNa8ZV02NseMIx0KssaAn9NkYFw4Y3Q5sYh3HQZsDRfja2HdFrnpIfQeTf3fjjyPWXw9iNRuPQHUXD0J1Hi3AZ3ywXy1n8822PN1KSi2PJCqrtWlO2x//KI+QzaqAprikIdXbPVrXX75r+AlBLAwQUAAAAAADZoXZHAAAAAAAAAAAAAAAADgAAAHhsL3dvcmtzaGVldHMvUEsDBBQAAAAIAAAAIQDmVajjXQEAAIQCAAAYAAAAeGwvd29ya3NoZWV0cy9zaGVldDEueG1sjZJPawIxEMXvhX6HkLtGbW2ruEpBpB4Kpf/u2ezsbjDJLMlY9dt3dq1S8OJtXibz471JZou9d+IHYrIYMjnsD6SAYLCwocrk1+eq9yRFIh0K7TBAJg+Q5GJ+ezPbYdykGoAEE0LKZE3UTJVKpgavUx8bCNwpMXpNLGOlUhNBF92Qd2o0GDwor22QR8I0XsPAsrQGlmi2HgIdIRGcJvafatukE82ba3Bex8226Rn0DSNy6ywdOqgU3kzXVcCoc8e598N7bU7sTlzgvTURE5bUZ9yf0cvMEzVRTJrPCssJ2rWLCGUmn4dSzWfdxW8Lu/SvFqTzD3BgCAp+Iyna3eeIm7a55qNBO6ouZldd0LcoCij11tE77l7AVjUxZMxZ2hTT4rCEZHiXjOmPxmcTS02a60ZX8KpjZUMSDsru1qMU8YjpasKmqxiZIxH6k6o5OcRW3UlRItJJtG7P/2f+C1BLAwQUAAAACAAAACEApFPFz0EBAAAIBAAAEwAAAFtDb250ZW50X1R5cGVzXS54bWytk89OAjEQxu8mvkPTK9kWPBhjWDj456gc8AFqO8s2dNumUxDe3tmCHgiKBC/b7M583+/bdjqebjrH1pDQBl/zkRhyBl4HY/2i5m/z5+qOM8zKG+WCh5pvAfl0cn01nm8jICO1x5q3Ocd7KVG30CkUIYKnShNSpzK9poWMSi/VAuTNcHgrdfAZfK5y78En40do1Mpl9rShz7skCRxy9rBr7Fk1VzE6q1Wmulx7c0Cp9gRBytKDrY04oAYujxL6ys+Ave6VtiZZA2ymUn5RHXXJjZMfIS3fQ1iK302OpAxNYzWYoFcdSQTGBMpgC5A7J8oqOmX94DS/NKMsy+ifg3z7n8iR6bxh97w8QrE5AcS8dYAXow62vZj+RibhLIWINLkJzqd/jWavriIZQcr2j0SyPh948LvQT70Bc4Qtyz2efAJQSwECFAAUAAAACAAAACEAtVUwI+wAAABMAgAACwAAAAAAAAABAAAAAAAAAAAAX3JlbHMvLnJlbHNQSwECFAAUAAAACAAAACEA3kEW2XsBAAARAwAAEAAAAAAAAAABAAAAAAAVAQAAZG9jUHJvcHMvYXBwLnhtbFBLAQIUABQAAAAIAOehdkc+qGWw1QAAAG0BAAARAAAAAAAAAAEAIAAAAL4CAABkb2NQcm9wcy9jb3JlLnhtbFBLAQIUABQAAAAAANmhdkcAAAAAAAAAAAAAAAAJAAAAAAAAAAAAEAAAAMIDAAB4bC9fcmVscy9QSwECFAAUAAAACAAAACEAjYfacNoAAAAtAgAAGgAAAAAAAAABAAAAAADpAwAAeGwvX3JlbHMvd29ya2Jvb2sueG1sLnJlbHNQSwECFAAUAAAACAAAACEA3iPy024CAACxBQAADQAAAAAAAAABAAAAAAD7BAAAeGwvc3R5bGVzLnhtbFBLAQIUABQAAAAAANmhdkcAAAAAAAAAAAAAAAAJAAAAAAAAAAAAEAAAAJQHAAB4bC90aGVtZS9QSwECFAAUAAAACAAAACEAi4JuWPUFAACOGgAAEwAAAAAAAAABAAAAAAC7BwAAeGwvdGhlbWUvdGhlbWUxLnhtbFBLAQIUABQAAAAIAAAAIQB8PO7DLgIAAJsEAAAPAAAAAAAAAAEAAAAAAOENAAB4bC93b3JrYm9vay54bWxQSwECFAAUAAAAAADZoXZHAAAAAAAAAAAAAAAADgAAAAAAAAAAABAAAAA8EAAAeGwvd29ya3NoZWV0cy9QSwECFAAUAAAACAAAACEA5lWo410BAACEAgAAGAAAAAAAAAABAAAAAABoEAAAeGwvd29ya3NoZWV0cy9zaGVldDEueG1sUEsBAhQAFAAAAAgAAAAhAKRTxc9BAQAACAQAABMAAAAAAAAAAQAAAAAA+xEAAFtDb250ZW50X1R5cGVzXS54bWxQSwUGAAAAAAwADADoAgAAbRMAAAAA',
        'base64'
    );
    this._initialize(data);
};

/**
 * Initialize the workbook. (This is separated from the constructor to ease testing.)
 * @param {Buffer} data - File buffer of the Excel workbook.
 * @returns {undefined}
 * @private
 */
Workbook.prototype._initialize = function (data) {
    this._zip = new JSZip(data, { base64: false, checkCRC32: true });
    var workbookText = this._zip.file("xl/workbook.xml").asText();
    this._workbookXML = parser.parseFromString(workbookText).documentElement;

    var relsText = this._zip.file("xl/_rels/workbook.xml.rels").asText();
    this._relsXML = parser.parseFromString(relsText).documentElement;

    this._sheets = [];
    this._sheetsNode = xpath("sml:sheets", this._workbookXML)[0];

    var sheetNodes = this._sheetsNode.childNodes;
    for (var i = 0; i < sheetNodes.length; i++) {
        var sheetText = this._zip.file("xl/worksheets/sheet" + (i + 1) + ".xml").asText();
        var sheetXML = parser.parseFromString(sheetText).documentElement;

        // This is a blunt way to make sure formula values get updated.
        // It just clears all stored values in case the referenced cell values change.
        var valueNodes = xpath("sml:sheetData/sml:row/sml:c/sml:f/../sml:*[name(.) !='f']", sheetXML);
        valueNodes.forEach(function (valueNode) {
            valueNode.parentNode.removeChild(valueNode);
        });

        var sheet = new Sheet(this, sheetNodes[i], sheetXML);
        this._sheets.push(sheet);
    }
};

/**
 * Create a new sheet.
 * @param {string} sheetName - The name of the sheet. Must be unique.
 * @param {number} [index] - The position of the sheet (0-based). Omit to insert at the end.
 * @returns {Sheet} The new sheet.
 */
Workbook.prototype.createSheet = function (sheetName, index) {
    if (index === undefined) index = this._sheets.length;
    if (!utils.isInteger(index) || index < 0 || index > this._sheets.length) {
        throw new Error("Invalid sheet index.");
    }

    // Create the new XML nodes.
    var sheetXML = parser.parseFromString('<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData/></worksheet>').documentElement;
    var sheetNode = parser.parseFromString('<sheet name="' + sheetName + '"/>').documentElement;

    // Insert the sheet definition node in the right place.
    if (index === this._sheets.length) {
        this._sheetsNode.appendChild(sheetNode);
    } else {
        this._sheetsNode.insertBefore(sheetNode, this._sheetsNode.childNodes[index]);
    }

    // Clear all the old sheet rel nodes.
    for (var i = this._relsXML.childNodes.length - 1; i >= 0; i--) {
        var rnode = this._relsXML.childNodes[i];
        if (rnode.getAttribute("Type") === "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet") {
            this._relsXML.removeChild(rnode);
        }
    }

    // Fix the sheet IDs to match the order.
    for (var j = 0; j < this._sheetsNode.childNodes.length; j++) {
        var id = j + 1;
        var snode = this._sheetsNode.childNodes[j];
        snode.setAttribute("sheetId", id);
        snode.setAttribute("r:id", "xpopId" + id);

        // Create a new sheet rel node.
        var relNode = parser.parseFromString('<Relationship Id="xpopId' + id + '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet' + id + '.xml"/>');
        this._relsXML.appendChild(relNode);
    }

    // Create the sheet and store it.
    var sheet = new Sheet(this, sheetNode, sheetXML);
    this._sheets.splice(index, 0, sheet);
    return sheet;
};

/**
 * Gets the sheet with the provided name or index (0-based).
 * @param {string|number} sheetNameOrIndex - The sheet name or index.
 * @returns {Sheet} The sheet, if found.
 */
Workbook.prototype.getSheet = function (sheetNameOrIndex) {
    if (utils.isInteger(sheetNameOrIndex)) return this._sheets[sheetNameOrIndex];

    for (var i = 0; i < this._sheets.length; i++) {
        var sheet = this._sheets[i];
        if (sheet.getName() === sheetNameOrIndex) return sheet;
    }
};

/**
 * Get a named cell. (Assumes names with workbook scope pointing to single cells.)
 * @param {string} cellName - The name of the cell.
 * @returns {Cell} The cell, if found.
 */
Workbook.prototype.getNamedCell = function (cellName) {
    var definedName = xpath("sml:definedNames/sml:definedName[@name='" + cellName + "']", this._workbookXML)[0];
    if (!definedName) return;

    var address = definedName.firstChild.nodeValue;
    var ref = utils.addressToRowAndColumn(address);
    if (!ref) return;

    return this.getSheet(ref.sheet).getCell(ref.row, ref.column);
};

/**
 * Gets the output.
 * @param {Object} [options] - The options for JSZip generate.
 * @returns {Buffer} A node buffer for the generated Excel workbook.
 */
Workbook.prototype.output = function (options) {
    options = options || {
        'type': 'nodebuffer'
    };

    this._zip.file("xl/workbook.xml", this._workbookXML.toString());
    this._zip.file("xl/_rels/workbook.xml.rels", this._relsXML.toString());

    for (var i = 0; i < this._sheets.length; i++) {
        var index = i + 1;
        var sheet = this._sheets[i];
        this._zip.file("xl/worksheets/sheet" + index + ".xml", sheet._sheetXML.toString());
    }

    // Kill the calc chain since this will corrupt the file is formulas are removed.
    this._zip.remove("xl/calcChain.xml");

    return this._zip.generate(options);
};

/**
 * Writes to file with the given path.
 * @param {string} path - The path of the file.
 * @param {function} cb - A callback.
 * @returns {undefined}
 */
Workbook.prototype.toFile = function (path, cb) {
    fs.writeFile(path, this.output(), cb);
};

/**
 * Writes to file with the given path synchronously.
 * @param {string} path - The path of the file.
 * @returns {undefined}
 */
Workbook.prototype.toFileSync = function (path) {
    fs.writeFileSync(path, this.output());
};

/**
 * Generates javascript blob object, to be used for client-side.
 * returns {Blob}
 */
Workbook.prototype.toBlob = function () {
    return this.output({
        'type': 'blob',
        'mimeType': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    });
};

/**
 * Creates a Workbook from the file with the given path.
 * @param {string} path - The path of the file.
 * @param {function} cb - A callback with the new workbook.
 * @returns {undefined}
 */
Workbook.fromFile = function (path, cb) {
    var x = function(err, data) {
        cb(err, new Workbook(data));
    };
    if (utils.isBrowser()) {
        var JSZipUtils = require('jszip-utils');
        JSZipUtils.getBinaryContent(path, x);
    }
    else {
        fs.readFile(path, x);
    }
};

/**
 * Creates a Workbook from the file with the given path synchronously.
 * @param {string} path - The path of the file.
 * @returns {Workbook} The parsed workbook.
 */
Workbook.fromFileSync = function (path) {
    if (utils.isBrowser()) {
        throw new Error('fromFileSync is unavailable to client-side applications');
    }
    else {
        var data = fs.readFileSync(path);
        return new Workbook(data);
    }
};

module.exports = Workbook;
