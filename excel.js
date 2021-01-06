let nevCalcTable = {
    EUP: [
        {
            "label": "Parma",
            "key": "8400149083",
            "value": "4.5"
        }, {
            "label": "Ferrara",
            "key": "8400150707",
            "value": "4.5"
        }, {
            "label": "Firenze",
            "key": "8400141787",
            "value": "4.5"
        }, {
            "label": "Modena-Reggio 1",
            "key": "8400124337",
            "value": "4.5"
        }, {
            "label": "Rovigo 1",
            "key": "8400118979",
            "value": "4.5"
        }, {
            "label": "Vicenza",
            "key": "8400141790",
            "value": "4.5"
        }, {
            "label": "Mantova-Cremona",
            "key": "8400149736",
            "value": "4.5"
        }, {
            "label": "Padova-Rovigo 2",
            "key": "8400149816",
            "value": "4.5"
        }, {
            "label": "Reggio-Modena 2",
            "key": "8400151041",
            "value": "4.5"
        }],
    CEP: []
}


let selectedFile;
//console.log(window.XLSX);
document.getElementById('input').addEventListener("change", (event) => {
    selectedFile = event.target.files[0];
})

let data = [{}];

document.getElementById('button').addEventListener("click", () => {
    XLSX.utils.json_to_sheet(data, 'out.xlsx');
    if (selectedFile) {
        let fileReader = new FileReader();
        fileReader.readAsBinaryString(selectedFile);

        fileReader.onload = (event) => {
            let data = event.target.result;
            let workbook = XLSX.read(data, { type: "binary" });
            workbook.SheetNames.forEach(sheet => {
                let rowObject = XLSX.utils.sheet_to_row_object_array(workbook.Sheets[sheet]);
                Calc(rowObject);
            });
        }
    }
});

let resulLoad = [{}]
function Calc(data) {
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
            "TOT": 0,
            "CON": 0,
            "ANN": 0,
            "AV": 0
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
                        LCLs[j].TOT += 1;
                        if (data[i]["Stato OdL"].localeCompare("Annullato") == 0) {
                            LCLs[j].ANN += 1;
                        } else if (data[i]["Stato OdL"].localeCompare("Chiuso") == 0 && (data[i]["Causale Esito"].localeCompare("OK FINALE") == 0 || data[i]["Causale Esito"].localeCompare("CHIUSO DA BACK OFFICE") == 0)) {
                            LCLs[j].CON += 1;
                        } else if (data[i]["Stato OdL"].localeCompare("Chiuso") == 0 && data[i]["Causale Esito"].localeCompare("Chiusura Giornata Lavorativa") != 0) {
                            LCLs[j].AV += 1;
                        }
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
                        "TOT": 0,
                        "CON": 0,
                        "ANN": 0,
                        "AV": 0
                    };

                    LCL.TOT += 1;
                    if (data[i]["Stato OdL"].localeCompare("Annullato") == 0) {
                        LCL.ANN += 1;
                    } else if (data[i]["Stato OdL"].localeCompare("Chiuso") == 0 && (data[i]["Causale Esito"].localeCompare("OK FINALE") == 0 || data[i]["Causale Esito"].localeCompare("CHIUSO DA BACK OFFICE") == 0)) {
                        LCL.CON += 1;
                    } else if (data[i]["Stato OdL"].localeCompare("Chiuso") == 0 && data[i]["Causale Esito"].localeCompare("Chiusura Giornata Lavorativa") != 0) {
                        LCL.AV += 1;
                    }

                    LCLs.push(LCL);

                }
            }
        }
        //console.log(LCLs);

        loadData(LCLs);
        resulLoad = LCLs;
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
                element.id = data[i].LCL;

                var typeLCL = "NaN";
                switch (data[i].TYPE) {
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

                element.innerHTML = '<b>' + data[i].LCL + '</b><i class="w3-tiny"> (' + data[i].CN + ', ' + typeLCL + ')</i><span onclick="changeCN(this.parentElement)" class="w3-button w3-transparent w3-display-right">&times;</span>';
                document.querySelector("#addListLCL").appendChild(element);
            }
        }

        document.querySelector("#loadFile").style.display = "none";
        document.querySelector("#selectLCL").style.display = "block";
    }
}

window.calcBeneficit = function () {
    var LCList = document.querySelector("#addListLCL").children;
    //console.log(LCList);

    document.querySelector("#selectLCL").style.display = "none";
    document.querySelector("#BeneficitTab").style.display = "block";
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

window.options = function () {
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

    var element = document.createElement("ul");
    element.classList.add("w3-ul");
    element.classList.add("w3-card-4");
    element.classList.add("w3-margin-top");
    element.classList.add("w3-margin-bottom");
    element.innerHTML = '<!-- Injection JavaScript --><li><h2>€/Punto</h2></li>';

    for (let i = 0; i < jsonCalcTable.EUP.length; i++) {
        element.innerHTML += '<li class="w3-display-container" id="' + jsonCalcTable.EUP[i].key + '"><b>' + jsonCalcTable.EUP[i].label + '</b><i class="w3-tiny">(' + jsonCalcTable.EUP[i].key + ')</i><span title="Edit" onclick="editData(this.parentElement);" class="w3-button w3-transparent w3-display-right w3-hover-dark-grey">' + jsonCalcTable.EUP[i].value + '€</span></li>';
    }

    document.querySelector("#optionsList").appendChild(element);

    document.querySelector('#selectLCL').style.display = 'none';
    document.querySelector('#optionsTab').style.display = 'block';
}

const elementID = document.querySelector('#ModalButtonSave');
function editData(element) {
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

            loadCalcTable.EUP[i].value = document.querySelector('#euroPunto').value;
            localStorage.setItem("calcTable", JSON.stringify(loadCalcTable));

            document.getElementById('modalEditOp').style.display = "none";

            document.querySelector("#optionsList").innerHTML = "<!-- Injection JavaScript -->";
            options();
        }
    }

    elementID.removeEventListener('click', saveCalcTable);
}

function closeModal() {
    document.getElementById('modalEditOp').style.display='none';
    elementID.removeEventListener('click', saveCalcTable);
}