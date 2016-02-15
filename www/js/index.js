/*
 * Licensed to the Apache Software Foundation (ASF) under one
 * or more contributor license agreements.  See the NOTICE file
 * distributed with this work for additional information
 * regarding copyright ownership.  The ASF licenses this file
 * to you under the Apache License, Version 2.0 (the
 * "License"); you may not use this file except in compliance
 * with the License.  You may obtain a copy of the License at
 *
 * http://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing,
 * software distributed under the License is distributed on an
 * "AS IS" BASIS, WITHOUT WARRANTIES OR CONDITIONS OF ANY
 * KIND, either express or implied.  See the License for the
 * specific language governing permissions and limitations
 * under the License.
 */
var X = XLSX;
var XW = {
    /* worker message */
    msg: 'xlsx',
    /* worker scripts */
    rABS: './xlsxworker2.js',
    norABS: './xlsxworker1.js',
    noxfer: './xlsxworker.js'
};
var SHEET_NAMES = ["survey", "choices", "settings"];

var SURVEY_SHEET_COLUMNS = ["type",
    "name",
    "label",
    "hint",
    "constraint",
    "constraintMessage",
    "required",
    "appearance",
    "default",
    "relevant",
    "readOnly",
    "calculation",
    "image",
    "audio"];
var CHOICES_SHEET_COLUMNS = ["name", "label", "image"];

var app = {
    // Application Constructor
    initialize: function () {
        this.bindEvents();
    },
    // Bind Event Listeners
    //
    // Bind any events that are required on startup. Common events are:
    // 'load', 'deviceready', 'offline', and 'online'.
    bindEvents: function () {
        document.addEventListener('deviceready', this.onDeviceReady, false);
        document.addEventListener('load', this.onLoad, false);
        document.addEventListener('offline', this.onOffline, false);
        document.addEventListener('online', this.onOnline, false);
    },
    // deviceready Event Handler
    //
    // The scope of 'this' is the event. In order to call the 'receivedEvent'
    // function, we must explicitly call 'app.receivedEvent(...);'
    onDeviceReady: function () {
        app.receivedEvent('deviceready');
    },
    onLoad: function () {
        app.receivedEvent('load');
    },
    onOffline: function () {
        app.receivedEvent('offline');
    },
    onOnline: function () {
        app.receivedEvent('online');
    },
    // Update DOM on a Received Event
    receivedEvent: function (id) {
        /*  var parentElement = document.getElementById(id);
         var listeningElement = parentElement.querySelector('.listening');
         var receivedElement = parentElement.querySelector('.received');

         listeningElement.setAttribute('style', 'display:none;');
         receivedElement.setAttribute('style', 'display:block;');
         */

        var xlf = document.getElementById('xlsFiles');

        function handleFile(e) {

            var files = e.target.files;
            var f = files[0];
            {
                var reader = new FileReader();
                var name = f.name;
                reader.onload = function (e) {
                    if (typeof console !== 'undefined')
                        console.log("onload", new Date());
                    var data = e.target.result;
                    if (use_worker) {
                        xw(data, process_wb);
                    } else {
                        var wb;
                        if (rABS) {
                            wb = X.read(data, {type: 'binary'});
                        } else {
                            var arr = fixdata(data);
                            wb = X.read(btoa(arr), {type: 'base64'});
                        }
                        process_wb(wb);
                    }
                };
                if (rABS) reader.readAsBinaryString(f);
                else reader.readAsArrayBuffer(f);
            }
        }

        if (xlf.addEventListener) xlf.addEventListener('change', handleFile, false);

        console.log('Received Event: ' + id);
    }
};

var rABS = typeof FileReader !== "undefined" && typeof FileReader.prototype !== "undefined" && typeof FileReader.prototype.readAsBinaryString !== "undefined";

var use_worker = typeof Worker !== 'undefined';
use_worker = false;

var transferable = use_worker;
transferable = false;


function xw_noxfer(data, cb) {
    var worker = new Worker(XW.noxfer);
    worker.onmessage = function (e) {
        switch (e.data.t) {
            case 'ready':
                break;
            case 'e':
                console.error(e.data.d);
                break;
            case XW.msg:
                cb(JSON.parse(e.data.d));
                break;
        }
    };
    var arr = rABS ? data : btoa(fixdata(data));
    worker.postMessage({d: arr, b: rABS});
}

function xw_xfer(data, cb) {
    var worker = new Worker(rABS ? XW.rABS : XW.norABS);
    worker.onmessage = function (e) {
        switch (e.data.t) {
            case 'ready':
                break;
            case 'e':
                console.error(e.data.d);
                break;
            default:
                xx = ab2str(e.data).replace(/\n/g, "\\n").replace(/\r/g, "\\r");
                console.log("done");
                cb(JSON.parse(xx));
                break;
        }
    };
    if (rABS) {
        var val = s2ab(data);
        worker.postMessage(val[1], [val[1]]);
    } else {
        worker.postMessage(data, [data]);
    }
}

function xw(data, cb) {

    if (transferable) xw_xfer(data, cb);
    else xw_noxfer(data, cb);
}

function process_wb(wb) {

    for (i = 0; i < SHEET_NAMES.length; i++) {
        if (SHEET_NAMES[i] != wb.SheetNames[i]) {
            if (typeof console !== 'undefined')
                console.log("sheet names error !!! expected [" + SHEET_NAMES[i] + "] found [" + wb.SheetNames[i] + "]");
            //todo alert message to user by updating DOM
            return;
        }
    }


    output = process_sheets(wb);

    if (out.innerText === undefined) out.textContent = output;
    else out.innerText = output;

    if (typeof console !== 'undefined')
        console.log("output", new Date());
}

function process_sheets(workbook) {
    var result = [];
    /*    workbook.SheetNames.forEach(function (sheetName) {
     var csv = sheet_to_csv(workbook.Sheets[sheetName]);
     if (csv.length > 0) {
     result.push("SHEET: " + sheetName);
     result.push("");
     result.push(csv);
     }
     });*/
    process_choices(workbook.Sheets["choices"]);
    process_survey(workbook.Sheets['survey']);
    return result.join("\n");
}

function process_choices(sheet, opts) {
    var out = "", txt = "", qreg = /"/g;

    if (sheet == null || sheet["!ref"] == null) return "";
    var r = X.utils.decode_range(sheet["!ref"])
    var FS = ",";
    var RS = "\n";
    var row = "", rr = "", cols = [];
    var i = 0, cc = 0, val;
    var R = 0, C = 0;
    for (C = r.s.c; C <= r.e.c; ++C) {
        cols[C] = X.utils.encode_col(C);
    }
    var lastListName = '';
    var items = [];
    for (R = r.s.r + 1; R <= r.e.r; ++R) {

        var choiceItem = new ChoiceItem();
        rr = X.utils.encode_row(R);
        for (C = r.s.c; C <= r.e.c; ++C) {
            val = sheet[cols[C] + rr];
            //txt = val !== undefined ? '' + format_cell(val) : "";
            if (C == 0 && val != undefined) {// if is first column means list name.
                lastListName = val.v;
            } else {
                if (val !== undefined) {
                    var s = 'set' + capitalizeFirstLetter(CHOICES_SHEET_COLUMNS[C - 1]);
                    choiceItem[s](val.v);
                } else {
                    if (C == 0) {
                        var choice = new Choice();
                        choice.setListName(lastListName);
                        choice.setItems(items);
                        choiceList.push(choice);
                        var items = [];
                        break;
                    }
                    console.log('next one...');
                }
            }


        }
        if (choiceItem.name != undefined)
            items.push(choiceItem);

    }
    return out;
}
function process_survey(sheet, opts) {
    var out = "", txt = "", qreg = /"/g;

    if (sheet == null || sheet["!ref"] == null) return "";
    var r = X.utils.decode_range(sheet["!ref"])
    var FS = ",";
    var RS = "\n";
    var row = "", rr = "", cols = [];
    var i = 0, cc = 0, val;
    var R = 0, C = 0;
    for (C = r.s.c; C <= r.e.c; ++C) {
        cols[C] = X.utils.encode_col(C);
    }

    for (R = r.s.r; R <= r.e.r; ++R) {
        row1 = "";
        row = new Row();
        rr = X.utils.encode_row(R);
        for (C = r.s.c; C <= r.e.c; ++C) {
            val = sheet[cols[C] + rr];
            //txt = val !== undefined ? '' + format_cell(val) : "";

            var s = 'set' + capitalizeFirstLetter(SURVEY_SHEET_COLUMNS[C]);
            if (val !== undefined) {
                row[s](val.v);
            } else {
                if (C == 0) {
                    break;
                }
                console.log('hoy');
            }
            //row1 += (C === r.s.c ? "" : FS) + txt;

        }
        out += row1 + RS;
        if (row.type != undefined)
            rowList.push(row);
    }
    return out;
}

function capitalizeFirstLetter(string) {
    return string[0].toUpperCase() + string.slice(1);
}


/***********

 MetaInfo Holder

 **********/

var rowList = [];
var choiceList = [];

function Row() {
}
Row.prototype = {
    constructor: Row,

    setType: function (type) {
        this.type = type;
    },
    setName: function (name) {
        this.name = name;
    },
    setLabel: function (label) {
        this.label = label;
    },
    setHint: function (hint) {
        this.hint = hint;
    },
    setConstraint: function (constraint) {
        this.constraint = constraint;
    },
    setConstraintMessage: function (constraintMessage) {
        this.constraintMessage = constraintMessage;
    },
    setRequired: function (required) {
        this.required = required;
    },
    setAppearance: function (appearance) {
        this.appearance = appearance;
    },
    setDefault: function (_default) {
        this._default = _default;
    },
    setRelevant: function (relevant) {
        this.relevant = relevant;
    },
    setReadOnly: function (readOnly) {
        this.readOnly = readOnly;
    },
    setCalculation: function (calculation) {
        this.calculation = calculation;
    },
    setImage: function (image) {
        this.image = image;
    },
    setAudio: function (audio) {
        this.audio = audio;
    },
    setChoice: function (choice) {
        this.choice = choice;
    }
};


function Choice() {
    this.items = [];
}
Choice.prototype = {
    constructor: Choice,

    setListName: function (listName) {
        this.listName = listName;
    },

    setItems: function (items) {
        this.items = items;
    },

    addItem: function (item) {
        this.items.push(item);
    }
};
function ChoiceItem() {

}
ChoiceItem.prototype = {
    constructor: ChoiceItem,

    setName: function (name) {
        this.name = name;
    },
    setLabel: function (label) {
        this.label = label;
    },
    setImage: function (image) {
        this.image = image;
    }
};

app.initialize();
