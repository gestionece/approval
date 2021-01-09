let nevCalcTable = {
    EUP: [
        {
            "label": "Parma",
            "key": "8400149083",
            "value": 4.5
        }, {
            "label": "Ferrara",
            "key": "8400150707",
            "value": 4.5
        }, {
            "label": "Firenze",
            "key": "8400141787",
            "value": 4.5
        }, {
            "label": "Modena-Reggio 1",
            "key": "8400124337",
            "value": "4.5"
        }, {
            "label": "Rovigo 1",
            "key": "8400118979",
            "value": 4.5
        }, {
            "label": "Vicenza",
            "key": "8400141790",
            "value": 4.5
        }, {
            "label": "Mantova-Cremona",
            "key": "8400149736",
            "value": 4.5
        }, {
            "label": "Padova-Rovigo 2",
            "key": "8400149816",
            "value": 4.5
        }, {
            "label": "Reggio-Modena 2",
            "key": "8400151041",
            "value": 4.5
        }],
    CEP: [
        {
            "label": "Eseguiti MF-TF",
            "filter": "MF",
            "key": "CON",
            "value": 2.5
        }, {
            "label": "Acesso a Vuoto MF-TF",
            "filter": "MF",
            "key": "AV",
            "value": 2
        }, {
            "label": "Eseguiti TF15/30",
            "filter": "TF",
            "key": "CON",
            "value": 3.5
        }, {
            "label": "Acesso a Vuoto TF15/30",
            "filter": "TF",
            "key": "AV",
            "value": 3
        }, {
            "label": "Eseguiti M2",
            "filter": "M2",
            "key": "CON",
            "value": 3.5
        }, {
            "label": "Acesso a Vuoto M2",
            "filter": "M2",
            "key": "AV",
            "value": 3
        }, {
            "label": "Recuperi Eseguiti MF-TF",
            "filter": "MF-R",
            "key": "CON",
            "value": 4.5
        }, {
            "label": "Recuperi Acesso a Vuoto MF-TF",
            "filter": "MF-R",
            "key": "AV",
            "value": 3.5
        }, {
            "label": "Recuperi Eseguiti TF15/30",
            "filter": "TF-R",
            "key": "CON",
            "value": 5.5
        }, {
            "label": "Recuperi Acesso a Vuoto TF15/30",
            "filter": "TF-R",
            "key": "AV",
            "value": 4.5
        }, { //GG MF-R
            "label": "Eseguiti <30 Giorni MF-TF",
            "filter": "MF-R",
            "key": "GG1",
            "value": 1
        }, {
            "label": "Eseguiti 30<>90 Giorni MF-TF",
            "filter": "MF-R",
            "key": "GG2",
            "value": 0.5
        }, {
            "label": "Eseguiti >90 Giorni MF-TF",
            "filter": "MF-R",
            "key": "GG3",
            "value": 0
        }, {
            "label": "Eseguiti <30 Giorni  TF15/30",
            "filter": "TF-R",
            "key": "GG1",
            "value": 1.5
        }, {
            "label": "Eseguiti 30<>90 Giorni TF15/30",
            "filter": "TF-R",
            "key": "GG2",
            "value": 1
        }, {
            "label": "Eseguiti >90 Giorni TF15/30",
            "filter": "TF-R",
            "key": "GG3",
            "value": 0.5
        }
    ]
}


let selectedFile;
//console.log(window.XLSX);
document.getElementById('input').addEventListener("change", (event) => {
    selectedFile = event.target.files[0];
})

let saveLoadFile;
document.getElementById('button').addEventListener("click", () => {
    let data = [{}];
    XLSX.utils.json_to_sheet(data, 'out.xlsx');
    if (selectedFile) {
        let fileReader = new FileReader();
        fileReader.readAsBinaryString(selectedFile);

        fileReader.onload = (event) => {
            let data = event.target.result;
            let workbook = XLSX.read(data, { type: "binary" });
            workbook.SheetNames.forEach(sheet => {
                let rowObject = XLSX.utils.sheet_to_row_object_array(workbook.Sheets[sheet]);

                if (sheet.localeCompare("Riepilogo attività completo") == 0) {
                    saveLoadFile = rowObject;
                    getLCL(rowObject);
                }
            });
        }
    }
});

function convertDate(stringDate) {
    const regex = /\d{2}\/\d{2}\/\d{4}/gm;
    var dateParts = stringDate.match(regex)[0].split("/");
    // month is 0-based, that's why we need dataParts[1] - 1
    return new Date(+dateParts[2], dateParts[1] - 1, +dateParts[0]);
}

let saveListLCL = [{}];
function getLCL(data) {
    if (data[0].LCL !== undefined) {
        var typelcl = "NaN";
        if (data[1]["Motivazione richiesta"].localeCompare("PRM2") == 0 && data[1]["CIT"].localeCompare("525") == 0) {
            typelcl = "M2";
        } else if (data[1]["Motivazione richiesta"].localeCompare("RI2G") == 0 && data[1]["CIT"].localeCompare("521") == 0) {
            typelcl = "MF-R";
        }
        else if (data[1]["Motivazione richiesta"].localeCompare("RI2G") == 0 && data[1]["CIT"].localeCompare("523") == 0) {
            typelcl = "TF-R";
        }
        else if (data[1]["Motivazione richiesta"].localeCompare("MA2G") == 0 && data[1]["CIT"].localeCompare("520") == 0) {
            typelcl = "MF";
        }
        else if (data[1]["Motivazione richiesta"].localeCompare("MA2G") == 0 && data[1]["CIT"].localeCompare("523") == 0) {
            typelcl = "TF";
        }

        let LCLs = [{
            "CN": data[1]["Codice contratto"],
            "LCL": data[1].LCL,
            "TYPE": typelcl,
            "DATE": convertDate(data[1]["Data creazione"]),
            "SELECT": true,
        }];

        var count = Object.keys(data).length;
        for (let i = 0; i < count; i++) {
            var LCLexist = false;
            var lastEneltel;
            if (lastEneltel !== data[i]["Eneltel"]) {
                lastEneltel = data[i]["Eneltel"];

                for (let j = 0; j < LCLs.length; j++) {
                    if (LCLs[j].LCL != undefined && data[i].LCL == LCLs[j].LCL) {
                        LCLexist = true;
                    }

                }
                if (LCLexist == false) {
                    LCLexist = false;

                    var typelcl = "NaN";
                    if (data[i]["Motivazione richiesta"].localeCompare("PRM2") == 0 && data[i]["CIT"].localeCompare("525") == 0) {
                        typelcl = "M2";
                    } else if (data[i]["Motivazione richiesta"].localeCompare("RI2G") == 0 && data[i]["CIT"].localeCompare("521") == 0) {
                        typelcl = "MF-R";
                    }
                    else if (data[i]["Motivazione richiesta"].localeCompare("RI2G") == 0 && data[i]["CIT"].localeCompare("523") == 0) {
                        typelcl = "TF-R";
                    }
                    else if (data[i]["Motivazione richiesta"].localeCompare("MA2G") == 0 && data[i]["CIT"].localeCompare("520") == 0) {
                        typelcl = "MF";
                    }
                    else if (data[i]["Motivazione richiesta"].localeCompare("MA2G") == 0 && data[i]["CIT"].localeCompare("523") == 0) {
                        typelcl = "TF";
                    }

                    let LCL = {
                        "CN": data[i]["Codice contratto"],
                        "LCL": data[i].LCL,
                        "TYPE": typelcl,
                        "DATE": convertDate(data[i]["Data creazione"]),
                        "SELECT": true,
                    };

                    LCLs.push(LCL);

                }
            }
        }
        //console.log(LCLs);
        saveListLCL = LCLs;
        loadData(LCLs);
    } else {
        //alert("Errato");
    }
}

function loadData(data) {
    if (data.length > 0) {
        for (let i = 0; i < data.length; i++) {

            data.sort(function (a, b) {
                return a.LCL - b.LCL;
            });

            var LCLexist = false;
            for (let j = 0; j < document.querySelector("#addListLCL").childElementCount; j++) {
                if (document.querySelector("#addListLCL").children[j].id == data[i].LCL) {
                    LCLexist = true;
                }
            }

            if (LCLexist == false) {
                var element = document.createElement("li");
                element.classList.add("w3-display-container");
                element.setAttribute("onclick", 'if(event.target === this) { modalEditLCL(this); }');
                element.id = data[i].LCL;

                var typeLCL = convertTYPE(data[i].TYPE);

                var jsonCalcTable = loadOptions();
                for (let cnI = 0; cnI < jsonCalcTable.EUP.length; cnI++) {
                    if (data[i].CN == jsonCalcTable.EUP[cnI].key) {
                        element.innerHTML = '<b>' + data[i].LCL + '</b><i class="w3-tiny"> (' + jsonCalcTable.EUP[cnI].label + ', ' + typeLCL + ')</i><span onclick="changeCN(this.parentElement)" class="w3-button w3-transparent w3-display-right">&times;</span>';
                    }
                }
                document.querySelector("#addListLCL").appendChild(element);
            }
        }

        document.querySelector("#loadFile").style.display = "none";
        document.querySelector("#selectLCL").style.display = "block";
    }
}

window.Print = function () {
    var resultList = document.querySelector("#BeneficitTab").innerHTML;

    var a = window.open('', '', 'width=733,height=454');
    a.document.open("text/html");
    a.document.write('<html><head><title>');
    a.document.write('Beneficit LCL');
    a.document.write('</title>');
    a.document.write('<link rel="stylesheet" href="https://www.w3schools.com/w3css/4/w3.css">');
    a.document.write('<style>body{opacity: 0}@media print { body{opacity: 1}}</style>');
    a.document.write('</head><body style="overflow: hidden;">');
    a.document.write(resultList);
    a.document.write("</body><!--FIX LIVE SERVER--></html>");

    a.document.close(); // necessary for IE >= 10
    a.focus(); // necessary for IE >= 10*/

    setTimeout(function () { a.print(); a.close(); }, 600);
}

function loadOptions() {
    var jsonCalcTable;
    if (localStorage.getItem("calcTable")) {
        jsonCalcTable = JSON.parse(localStorage.getItem("calcTable"));
    } else {
        localStorage.setItem("calcTable", JSON.stringify(nevCalcTable));
        jsonCalcTable = nevCalcTable;
    }

    jsonCalcTable.EUP.sort(function (a, b) {
        return a.label - b.label;
    });
    jsonCalcTable.CEP.sort(function (a, b) {
        return b.value - a.value;
    });

    return jsonCalcTable;
}

window.options = function () {
    var jsonCalcTable = loadOptions();

    var element = document.createElement("ul");
    element.classList.add("w3-ul");
    element.classList.add("w3-card-4");
    element.classList.add("w3-margin-top");
    element.classList.add("w3-margin-bottom");
    element.innerHTML = '<!-- Injection JavaScript --><li><h2>€/Punto</h2></li>';

    for (let i = 0; i < jsonCalcTable.EUP.length; i++) {
        element.innerHTML += '<li class="w3-display-container" id="' + jsonCalcTable.EUP[i].key + '"><b>' + jsonCalcTable.EUP[i].label + '</b><i class="w3-tiny">(' + jsonCalcTable.EUP[i].key + ')</i><span title="Edit" onclick="ediTable(this.parentElement);" class="w3-button w3-transparent w3-display-right w3-hover-yellow">' + parseFloat(jsonCalcTable.EUP[i].value).toFixed(1) + '<i class="w3-tiny">€</i></span></li>';
    }

    document.querySelector("#optionsList").appendChild(element);

    element = document.createElement("ul");
    element.classList.add("w3-ul");
    element.classList.add("w3-card-4");
    element.classList.add("w3-margin-top");
    element.classList.add("w3-margin-bottom");
    element.innerHTML = '<!-- Injection JavaScript --><li><h2>Contatore/Punto</h2></li>';

    for (let i = 0; i < jsonCalcTable.CEP.length; i++) {
        element.innerHTML += '<li class="w3-display-container"><b>' + jsonCalcTable.CEP[i].label + '</b><span title="Edit" class="w3-button w3-transparent w3-display-right w3-hover-yellow">' + parseFloat(jsonCalcTable.CEP[i].value).toFixed(1) + '<i class="w3-tiny">p</i></span></li>';
    }

    document.querySelector("#optionsList").appendChild(element);

    document.querySelector('#selectLCL').style.display = 'none';
    document.querySelector('#optionsTab').style.display = 'block';
}

const elementID = document.querySelector('#ModalButtonSave');
function ediTable(element) {
    if (localStorage.getItem("calcTable")) {
        loadCalcTable = JSON.parse(localStorage.getItem("calcTable"));

        for (let i = 0; i < loadCalcTable.EUP.length; i++) {
            if (loadCalcTable.EUP[i].key == element.id) {
                document.querySelector('#labelCN').innerHTML = loadCalcTable.EUP[i].label + '<i class="w3-small">(' + loadCalcTable.EUP[i].key + ')</i>';
                document.querySelector('#euroPunto').value = loadCalcTable.EUP[i].value;
            }
        }

        document.getElementById('modalEditOp').style.display = "block";

        elementID.addEventListener('click', saveCalcTable, false);
        elementID.myParam = element;
    }
}

function saveCalcTable(evt) {
    for (let i = 0; i < loadCalcTable.EUP.length; i++) {
        if (loadCalcTable.EUP[i].key == evt.currentTarget.myParam.id) {

            loadCalcTable.EUP[i].value = parseFloat(document.querySelector('#euroPunto').value);
            localStorage.setItem("calcTable", JSON.stringify(loadCalcTable));

            document.getElementById('modalEditOp').style.display = "none";

            document.querySelector("#optionsList").innerHTML = "<!-- Injection JavaScript -->";
            options();
        }
    }

    elementID.removeEventListener('click', saveCalcTable);
}

function closeModal() {
    document.querySelector('#labelCN').innerHTML = "<!-- Injection JavaScript -->";
    document.querySelector('#euroPunto').value = "";
    document.getElementById('modalEditOp').style.display = 'none';
    elementID.removeEventListener('click', saveCalcTable);
}

function dateToYMD(date) {
    var d = date.getDate();
    var m = date.getMonth() + 1; //Month from 0 to 11
    var y = date.getFullYear();
    return '' + y + '-' + (m <= 9 ? '0' + m : m) + '-' + (d <= 9 ? '0' + d : d);
}

const elementIDlcl = document.querySelector('#ModalButtonSaveLCL');
function modalEditLCL(element) {
    document.getElementById('modalEditLCL').style.display = "block";
    for (let i = 0; i < saveListLCL.length; i++) {
        if (saveListLCL[i].LCL == element.id) {
            document.querySelector('#labelLCL').innerHTML = saveListLCL[i].LCL + '<i class="w3-small">(' + saveListLCL[i].CN + ')</i>';
            document.querySelector('#dateLCL').value = dateToYMD(saveListLCL[i].DATE);
            document.querySelector('#typeLCL').value = saveListLCL[i].TYPE;

            elementIDlcl.addEventListener('click', saveCalcTableLCL, false);
            elementIDlcl.myParam = element;
        }
    }
}

function saveCalcTableLCL(evt) {
    for (let i = 0; i < saveListLCL.length; i++) {
        if (saveListLCL[i].LCL == evt.currentTarget.myParam.id) {
            saveListLCL[i].DATE = new Date(document.querySelector('#dateLCL').value);
            saveListLCL[i].TYPE = document.querySelector('#typeLCL').value;

            document.querySelector('#labelLCL').innerHTML = "<!-- Injection JavaScript -->";
            document.getElementById('modalEditLCL').style.display = 'none';

            var typeLCL = convertTYPE(saveListLCL[i].TYPE);
            evt.currentTarget.myParam.innerHTML = '<b>' + saveListLCL[i].LCL + '</b><i class="w3-tiny"> (' + saveListLCL[i].CN + ', ' + typeLCL + ')</i><span onclick="changeCN(this.parentElement)" class="w3-button w3-transparent w3-display-right">&times;</span>';
        }
    }

    elementID.removeEventListener('click', saveCalcTableLCL);
}

function closeModaLCL() {
    document.querySelector('#labelLCL').innerHTML = "<!-- Injection JavaScript -->";
    document.getElementById('modalEditLCL').style.display = 'none';
    elementID.removeEventListener('click', saveCalcTableLCL);
}

let saveResultBeneficit;
function calcBeneficit() {
    let LCLs = [];
    for (let i = 0; i < saveListLCL.length; i++) {
        if (saveListLCL[i].SELECT == true) {

            let LCL = {
                "CN": saveListLCL[i].CN,
                "LCL": saveListLCL[i].LCL,
                "TYPE": saveListLCL[i].TYPE,
                "DATE": saveListLCL[i].DATE,
                "TOT": 0,
                "CON": 0,
                "ANN": 0,
                "AV": 0,
                "GG1": 0,
                "GG2": 0,
                "GG3": 0,
            };

            for (let ii = 0; ii < saveLoadFile.length; ii++) {
                if (saveListLCL[i].LCL == saveLoadFile[ii].LCL) {
                    LCL.TOT += 1;
                    if (saveLoadFile[ii]["Stato OdL"].localeCompare("Annullato") == 0) {
                        LCL.ANN += 1;
                    } else if (saveLoadFile[ii]["Stato OdL"].localeCompare("Chiuso") == 0 && (saveLoadFile[ii]["Causale Esito"].localeCompare("OK FINALE") == 0 || saveLoadFile[ii]["Causale Esito"].localeCompare("CHIUSO DA BACK OFFICE") == 0)) {
                        LCL.CON += 1;

                        const diffTime = Math.abs(new Date(LCL.DATE) - convertDate(saveLoadFile[ii]["Data e ora fine esecuzione"]));
                        const diffDays = Math.ceil(diffTime / (1000 * 60 * 60 * 24)) - 1;
                        if (diffDays <= 30) {
                            LCL.GG1 += 1;
                        } else if (diffDays > 30 && diffDays <= 90) {
                            LCL.GG2 += 1;
                        } else {
                            LCL.GG3 += 1;
                        }

                    } else if (saveLoadFile[ii]["Stato OdL"].localeCompare("Chiuso") == 0 && saveLoadFile[ii]["Causale Esito"].localeCompare("Chiusura Giornata Lavorativa") != 0) {
                        LCL.AV += 1;
                    }

                }
            }

            LCLs.push(LCL);

            //CODE
            var divObject = document.createElement('div');
            divObject.classList.add("w3-containery");
            divObject.classList.add("w3-light-grey");
            divObject.classList.add("w3-card-4");

            var typeLCL = convertTYPE(saveListLCL[i].TYPE);

            var jsonCalcTable = loadOptions();
            for (let cnI = 0; cnI < jsonCalcTable.EUP.length; cnI++) {
                if (saveListLCL[i].CN == jsonCalcTable.EUP[cnI].key) {
                    divObject.innerHTML = '<h2>' + saveListLCL[i].LCL + '<i class="w3-small"> (' + jsonCalcTable.EUP[cnI].label + ', ' + typeLCL + ')</i></h2><table id="lclPerCent" class="w3-table-all w3-hoverable w3-margin-bottom"><thead><tr class="w3-green"><th style="width: 40%;">Causale</th><th class="w3-center">Contatori</th><th class="w3-center">Punti</th><th class="w3-center">€/Punto</th><th class="w3-center">€</th></tr></thead><!-- Injection JavaScript --></table>';
                }
            }

            var subTot = 0;
            var formatter = new Intl.NumberFormat('it-IT', { style: 'currency', currency: 'EUR' });
            for (let j = 0; j < jsonCalcTable.CEP.length; j++) {
                for (let jj = 0; jj < jsonCalcTable.EUP.length; jj++) {
                    if (saveListLCL[i].TYPE == jsonCalcTable.CEP[j].filter && saveListLCL[i].CN == jsonCalcTable.EUP[jj].key) {

                        var row = document.createElement("tr");
                        var numVar = 0;
                        switch (jsonCalcTable.CEP[j].key) {
                            case "CON":
                                numVar = LCL.CON;
                                break;
                            case "AV":
                                numVar = LCL.AV;
                                break;
                            case "GG1":
                                numVar = LCL.GG1;
                                break;
                            case "GG2":
                                numVar = LCL.GG2;
                                break;
                            case "GG3":
                                numVar = LCL.GG3;
                                break;
                            default:
                                break;
                        }
                        var tot = numVar * jsonCalcTable.CEP[j].value * jsonCalcTable.EUP[jj].value;
                        subTot += tot;
                        row.innerHTML = "<td>" + jsonCalcTable.CEP[j].label + "</td><td class='w3-center'>" + numVar + "</td><td class='w3-center'>" + jsonCalcTable.CEP[j].value + "</td><td class='w3-center'>" + jsonCalcTable.EUP[jj].value + "€" + "</td><td class='w3-center'>" + formatter.format(tot) + "</td>";
                        divObject.querySelector("#lclPerCent").appendChild(row);
                    }
                }
            }

            var row = document.createElement("tr");
            row.innerHTML = "<td>" + "Totale:" + "</td><td></td><td></td><td></td><td class='w3-center'>" + formatter.format(subTot) + "</td>";
            divObject.querySelector("#lclPerCent").appendChild(row);

            document.querySelector("#listCnLCL").appendChild(divObject);
        }
    }

    saveResultBeneficit = LCLs;

    document.querySelector("#selectLCL").style.display = "none";
    document.querySelector("#BeneficitTab").style.display = "block";
}

function convertTYPE(type) {
    var typeLCL = "NaN";
    switch (type) {
        case "M2":
            typeLCL = "M2";
            break;
        case "MF-R":
            typeLCL = "MF-TF Recuperi";
            break;
        case "TF-R":
            typeLCL = "TF-15/30 Recuperi";
            break;
        case "MF":
            typeLCL = "MF-TF";
            break;
        case "TF":
            typeLCL = "TF-15/30";
            break;
        default:
            break;
    }
    return typeLCL;
}

function download_csv(filename = "beneficit") {

    var csv = 'CalcBeneficit,,,,\n';
    for (let i = 0; i < saveResultBeneficit.length; i++) {
        //Start Header CSV
        var row = saveResultBeneficit[i].LCL + ',';
        var jsonCalcTable = loadOptions();
        for (let cnI = 0; cnI < jsonCalcTable.EUP.length; cnI++) {
            if (saveResultBeneficit[i].CN == jsonCalcTable.EUP[cnI].key) {
                row += jsonCalcTable.EUP[cnI].label + ',';
            }
        }
        var typeLCL = convertTYPE(saveListLCL[i].TYPE);
        row += typeLCL + ',,\n';
        row += 'Causale,Contatori,Punti,Euro/Punto,Euro\n';
        //Start Table CSV
        var subTot = 0;
        for (let j = 0; j < jsonCalcTable.CEP.length; j++) {
            for (let jj = 0; jj < jsonCalcTable.EUP.length; jj++) {
                if (saveListLCL[i].TYPE == jsonCalcTable.CEP[j].filter && saveListLCL[i].CN == jsonCalcTable.EUP[jj].key) {

                    var numVar = 0;
                    switch (jsonCalcTable.CEP[j].key) {
                        case "CON":
                            numVar = saveResultBeneficit[i].CON;
                            break;
                        case "AV":
                            numVar = saveResultBeneficit[i].AV;
                            break;
                        case "GG1":
                            numVar = saveResultBeneficit[i].GG1;
                            break;
                        case "GG2":
                            numVar = saveResultBeneficit[i].GG2;
                            break;
                        case "GG3":
                            numVar = saveResultBeneficit[i].GG3;
                            break;
                        default:
                            break;
                    }
                    var tot = numVar * jsonCalcTable.CEP[j].value * jsonCalcTable.EUP[jj].value;
                    subTot += tot;

                    //console.log(  (tot).toString().replace(".",",")   );

                    row += jsonCalcTable.CEP[j].label + ',' + numVar + ',' + jsonCalcTable.CEP[j].value + ',' + jsonCalcTable.EUP[jj].value + ',' + tot + '\n';
                }
            }
        }

        //console.log( (subTot).toString().replace(".",",") );

        row += 'Totale:,,,,' + subTot + '\n';
        csv += row;
        csv += ',,,,\n,,,,\n';
    }
    
    var a = document.createElement("a");
    var url = 'data:text/csv;charset=utf-8,' + encodeURI(csv);
    a.href = url;
    a.download = filename + '.csv';
    document.body.appendChild(a);
    a.click();
    setTimeout(function () {
        document.body.removeChild(a);
        window.URL.revokeObjectURL(url);
    }, 0);
}