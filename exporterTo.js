

class ExportTo {

    _Header = [];
    _Body = [];
    _Shift = "";
    _Option = {
        direction: "RTL",
        position: "middle",
        border: 0,
        frame: "void",
        rules: "non",
        font_size_header: 24,
        font_size_caption: 26,
        font_size_body: 24,
        font_weight_header: "bold",
        font_weight_caption: "bold",
        font_weight_body: "normal",
        font_color_header: "black",
        font_color_caption: "black",
        font_color_body: "black",
        background_header: "transparent",
        background_caption: "transparent",
        background_body: "transparent",
        text_align: "center",
        shift_width: "auto",
        shift_max_width: 200,
        image_height: 300,
    };
    _Data = [];
    _Image = "";
    _Collection = "";
    _Caption = "";
    _To = FileTypes.XLS;
    Name = "";
    Space = 0;
    HeaderArray = [];


    /** default constructor take  deafult Option 
         * option = {direction: "RTL",position: "middle",border: 0,frame: "void",rules: "non",font_size_header: 24,font_size_caption: 26,font_size_body: 24,
        font_weight_header: "bold",font_weight_caption: "bold",font_weight_body: "normal",font_color_header: "black",font_color_caption: "black",font_color_body: "black",
        background_header:"transparent",background_caption:"transparent",background_body: "transparent",text_align:"center",shift_width: "auto",  shift_max_width: 200, image_height: 300,
        */

    // constructor() {

    //     if (this._Option.position == "middle")
    //         this.Shift(screen.width === 1920 ? 9 : 4);
    // }

    constructor(fileType) {
            this._To=fileType;
        if (this._Option.position == "middle")
        {
                if(this._To==FileTypes.PDF)
                {
                    this.Shift(0);
                }else
                {
                    this.Shift(screen.width === 1920 ? 9 : 4);
                }
        }
    }


    /**
    * constructor init all component
    * use Export() to download 
    * @param {string} caption
    * @param {[]} header
    * @param {any} body
    * @param {string} image
    * @param {FileTypes} to
    * @param {{}} option
    */
    Init(caption = null, header = null, body = null, image = null, to = null, option = null) {
        this._Option = option;
        if (this._Option.position == "middle")
        this._Shift = "";
            this.Shift(screen.width === 1920 ? 9 : 4);


        this.Header(header);
        this.Caption(caption);
        this.Body(body);
        this.Image(image);
        this.To(to);
        this.Collection();

        return this;
    }



    /**
     * constructor init
     * ExportTo csv , pdf , xls , txt , json .... 
     * option = {direction: "RTL",position: "middle",border: 0,frame: "void",rules: "non",font_size_header: 24,font_size_caption: 26,font_size_body: 24,
        font_weight_header: "bold",font_weight_caption: "bold",font_weight_body: "normal",font_color_header: "black",font_color_caption: "black",font_color_body: "black",
        background_header:"transparent",background_caption:"transparent",background_body: "transparent",text_align:"center",shift_width: auto,   shift_max_width: 200, image_height: 300,
    }
     * @param {any} option
     */
    Init(option = this._Option) {

        this._Option = option;


        if (this._Option.position == "middle")
        {
            this._Shift="";
            this.Shift(screen.width === 1920 ? 9 : 4);
        }

        return this;
    }



    /**
     * custome option other way with put constructor
     * its useful when there one init(constructor) and us
ing multi way to generate option
     * @param {any} option
     */
    Option(option) {
        this._Option = option;
        this._Shift="";
        if (this._Option.position == "middle")
        {
                if(this._To==FileTypes.PDF)
                {
                    this.Shift(0);
                }else
                {
                    this.Shift(screen.width === 1920 ? 9 : 4);
                }
        }
        return this;
    }

    /**
     * give pad cells
     * @param {number} space
     */
    Shift(space) {
        this.Space = space;
        this._Shift = "";//reset
        for (var i = 0; i < space; i++) {
            this._Shift += `<td style="width:${this._Option.shift_width}px;max-width:${this._Option.shift_max_width}px"></td>`;
        }
        return this;
    }

    /**
     *  captin top of header like a title take wide all col
     * @param {string} caption
     */
    Caption(caption) {

        this._Caption += `<tr>`;

        this.InsertShift(Schema.Caption);

        this._Caption += `<td colspan="${this.HeaderArray.length}" style="text-align:${this._Option.text_align};font-size:${this._Option.font_size_caption}px;font-weight:${this._Option.font_weight_caption};background-color:${this._Option.background_caption};color:${this._Option.font_color_caption};">${caption}</td>`;

        this._Caption += `</tr>`;

        return this;
    }

    /**
     * header like columns of tabel
     * @param {[]} header
     */
    Header(header) {
        this.HeaderArray = header;
        this._Header += `<tr>`;

        this.InsertShift(Schema.Header);

        header.forEach((head, index) => {
            this._Header += `<td style="text-align:${this._Option.text_align};font-size:${this._Option.font_size_header}px;font-weight:${this._Option.font_weight_header};background-color:${this._Option.background_header};color:${this._Option.font_color_header};"> ${head} </td>`;
        });
        this._Header += `</tr>`;

        return this;
    }

    /**
     * body contains the row of data
     * defaul should init data befor body  or init data at scop of body
     * @param {[][]}  data
     * @param {JSON}  data
     * @param {any}  data
     */
    Body(data) {

        if (this._Data.length === 0)
            this._Data = data;


        this._Data.forEach((row, index) => {

            this._Body += `<tr>`;

            this.InsertShift(Schema.Body);

            row.forEach((cell, iter) => {
                this._Body += `<td style="vertical-align:middle;text-align:${this._Option.text_align};font-size:${this._Option.font_size_body}px;background-color:${this._Option.background_body};color:${this._Option.font_color_body};"> ${cell} </td>`;
            });

            this._Body += `</tr>`;

        });


        return this;
    }


    /**
     * init data as rows to contain inside body
     * @param {any} data
     */
    Data(data) {
        this._Data = data;
        return this;
    }

    /**
     * convert table (html) as data to body , it's should confirm body as default after this
     * @param {string} tableid
     */
    Table(tableid) {
        this._Data = tabletodata(tableid);;
        return this;
    }


    /**
     * insert image top of page like symbol
     * @param {string} src
     */
    Image(src) {

        this._Image += `<tr>`;

        for (var i = 0; i < this.Space + parseInt((this.HeaderArray.length / 2) - 0.871); i++) {
            this._Image += `<td style="width:${this._Option.shift_width}px;max-width:${this._Option.shift_max_width}px"></td>`;
        }

        this._Image += ` <td style="vertical-align:middle; text-align:center;height:${this._Option.image_height}px">
                      <img height=${this._Option.image_height}  src="${src}"></img> </td>`;
        this._Image += `</tr>`;

        return this;
    }

    /** collector for each caption , header , body , image ..... */
    Collection() {
        //Attribute definitions

        //  frame = void| above | below | hsides | lhs | rhs | vsides | box | border[CI]
        //                              This attribute specifies which sides of the frame surrounding a table will be visible.Possible values:
        //      void: No sides.This is the default value.
        //      above: The top side only.
        //      below: The bottom side only.
        //      hsides: The top and bottom sides only.
        //      vsides: The right and left sides only.
        //      lhs: The left-hand side only.
        //      rhs: The right - hand side only.
        //      box: All four sides.
        //      border: All four sides.

        //  rules = none | groups | rows | cols | all[CI]
        //                              This attribute specifies which rules will appear between cells within a table.The rendering of rules is user agent dependent.Possible values:
        //      none: No rules.This is the default value.
        //      groups: Rules will appear between row groups(see THEAD, TFOOT, and TBODY) and column groups(see COLGROUP and COL) only.
        //      rows: Rules will appear between rows only.
        //      cols: Rules will appear between columns only.
        //      all: Rules will appear between all rows and columns.

        //  border = pixels[CN]
        //                              This attributes specifies the width(in pixels only) of the frame around a table(see the Note below for more information about this attribute).

        this._Collection += `<table style="table-layout: fixed; -fs-table-paginate: paginate;"  frame="${this._Option.frame}" rules="${this._Option.rules}" border="${this._Option.border}">`;
        this._Collection += this._Image;
        this._Collection += this._Caption;
        this._Collection += this._Header;
        this._Collection += this._Body;
        this._Collection += `</table>`;
        return this;
    }

    /**
     * important to here set type to convert as { xls , pdf , csv....... }
     * @param {FileTypes} type
     * @param {string} name
     */
    To(type, name) {
        this._To = type;
        this.Name = name;
        return this;
    }


     /**
     * filename
     * @param {string} name
     */
    AsName(name) {
        this.Name = name;
        return this;
    }

    /** but at last of init as trigger to start download
     * return true */
    Export() {

        this.Collection();
        var todaysDate =new Date().getTime().toString();
        let a = document.createElement('a');
        a.style.display = 'none';
        a.setAttribute('target', '_blank');

        let _name = this.Name;
        let _table = this._Collection;

        switch (this._To) {

            case FileTypes.XLS:

                let dir = this._Option.direction == "RTL" ? "<x:DisplayRightToLeft/>" : "";
                let blobURL = (function () {
                    var uri = 'data:application/vnd.ms-excel;base64,'
                        , template = `<html  xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel" xmlns="http://www.w3.org/TR/REC-html40">
            <head><!--[if gte mso 9]><xml><x:ExcelWorkbook><x:ExcelWorksheets><x:ExcelWorksheet><x:Name>{worksheet}</x:Name>
                            <x:WorksheetOptions><x:DisplayGridlines/>${dir}</x:WorksheetOptions></x:ExcelWorksheet></x:ExcelWorksheets></x:ExcelWorkbook></xml>
                                <![endif]--><meta http-equiv="content-type" content="text/plain; charset=UTF-8"/></head><body>
                                          <br>{table}
                    </body></html>`
                        , base64 = function (s) { return window.btoa(unescape(encodeURIComponent(s))) }
                        , format = function (s, c) { return s.replace(/{(\w+)}/g, function (m, p) { return c[p]; }) }

                    var ctx = { worksheet: _name || 'Worksheet', table: _table }
                    var blob = new Blob([format(template, ctx)]);
                    var blobURL = window.URL.createObjectURL(blob);
                    return blobURL;

                })();


                a.download = _name + "_" + todaysDate + '.xls';
                a.href = blobURL;

                break;


            case FileTypes.CSV:
                this._Data.unshift(this.HeaderArray);
                var filename = _name + "_" + todaysDate + '.csv';
                var csv = this._Data.join('\n');// array2CSV(this._Data);
                var blob = new Blob(["\ufeff", csv]);
                var url = URL.createObjectURL(blob);
                a.setAttribute('download', filename);
                a.setAttribute('href', url);
                break;

            case FileTypes.TEXT:
                this._Data.unshift(this.HeaderArray);
                var _string = this._Data.join('\n');
                var filename = _name + "_" + todaysDate + '.txt';
                a.setAttribute('download', filename);
                a.setAttribute('href', 'data:text/plain;charset=utf-8,' + encodeURIComponent(_string));

                break;

            case FileTypes.PDF:

                var divContents = _table;
                var printWindow = window.open('', '', 'height=400,width=800');

                let temp = this._Option.position == "middle" ? `style="margin:auto; width:50%;"` : "";

                printWindow.document.write('<html  dir='+this._Option.direction+'  '+temp+'  ><head><title>' + _name + "_" + todaysDate + '</title>');
                printWindow.document.write('</head><body onload="print()">');
                printWindow.document.write(divContents);
                printWindow.document.write('</body></html>');
                printWindow.document.close();
                //printWindow.document.addEventListener('DOMContentLoaded', printWindow.print());
               // printWindow.print();

                //this._Data.unshift(this.HeaderArray);
                //var _string = this._Data;
                //var filename = _name + "_" + todaysDate + '.pdf';
                //a.setAttribute('download', filename);
                //a.setAttribute('href', 'data:application/pdf;charset=utf-8,' + encodeURIComponent(_string));

                return true;

            case FileTypes.JSON:
                this._Data.unshift(this.HeaderArray);
                var _string = JSON.stringify(this._Data);
                var filename = _name + "_" + todaysDate + '.json';
                a.setAttribute('download', filename);
                a.setAttribute('href', 'data:application/json;charset=utf-8,' + encodeURIComponent(_string));
                break;


            default: console.error("unkown filetype  To()  not using a to specific type of FileTypes"); return;
        }


        document.body.appendChild(a)
        a.click()
        document.body.removeChild(a)

        return true;
    }


    /**
     * easy to repeat pad to all file 
     * @param {Schema} schema
     */
    InsertShift(schema) {
        switch (schema) {
            case Schema.Caption:
                this._Caption += this._Shift;
                return;

            case Schema.Header:
                this._Header += this._Shift;
                return;

            case Schema.Body:
                this._Body += this._Shift;
            case Schema.Image:
                this._Image += this._Shift;
                return;
            default: return;
        }
    }


}

/** enum useful with shift */
const Schema = {
    Caption: 0,
    Header: 1,
    Body: 2,
    Image: 3,
}

/** using as enum */
const FileTypes = {
    XLS: 0,
    XLSX: 1,
    CSV: 2,
    TEXT: 3,
    PDF: 4,
    JSON: 5,
}


/**
 * convert table (html) to matrix
 * @param {string} table_id
 */
function tabletodata(table_id) {
    var rows = document.querySelectorAll('table#' + table_id + ' tr');
    var data = [];
    for (var i = 0; i < rows.length; i++) {
        var row = [], cols = rows[i].querySelectorAll('td, th');
        for (var j = 0; j < cols.length; j++) {
            var data = cols[j].innerText.replace(/(\r\n|\n|\r)/gm, '').replace(/(\s\s)/gm, ' ')
            data = data.replace(/"/g, '""');
            row.push('"' + data + '"');
        }
        data.push(row.join(';'));
    }
    return data;
}

/**
 * swap array to coma
 * @param {any} array
 */
function array2CSV(array) {
    //array.splice.apply(array, [0, 0].concat(c));
    var str = '';
    for (var i = 0; i < array.length; i++) {
        var line = '';
        for (var index in array[i]) {
            line += array[i][index] + ',';
        }
        line = line.slice(0, -1);
        str += line + '\r\n';
    }
    return str;
}


