let selectedFile;
console.log(window.XLSX);
document.getElementById('input').addEventListener("change", (event) => {
    selectedFile = event.target.files[0];
})

let data = [{}]

document.getElementById('button').addEventListener("click", () => {
    XLSX.utils.json_to_sheet(data, 'out.xlsx');
    if (selectedFile) {
        let fileReader = new FileReader();
        fileReader.readAsBinaryString(selectedFile);

        fileReader.onload = (event) => {
            let data = event.target.result;

            let workbook = XLSX.read(data, { type: "binary" });
            //console.log(workbook);
            workbook.SheetNames.forEach(sheet => {
                let rowObject = XLSX.utils.sheet_to_row_object_array(workbook.Sheets[sheet]);
                //console.log(rowObject);
                //document.getElementById("jsondata").innerHTML = JSON.stringify(rowObject, undefined, 4);
                Calc(rowObject);
            });
        }
    }
});

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
        console.log(LCLs);
        //document.getElementById("jsondata").innerHTML = JSON.stringify(LCLs, undefined, 4);

        loadData(LCLs);
    }
}

function loadData(data) {
    if (data.length > 0) {
        for (let i = 0; i < data.length; i++) {

            data.sort(function (a, b) {
                return a.LCL - b.LCL;
            });

            var LCLexist = false;
            for (let j = 0; j < document.querySelector("#addListCN").childElementCount; j++) {
                if (document.querySelector("#addListCN").children[j].id == data[i].LCL) {
                    LCLexist = true;
                }
            }

            if (LCLexist == false) {
                var element = document.createElement("li");
                element.classList.add("w3-display-container"); //<i class="w3-tiny"> (update 3 day ago)</i>
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

                element.innerHTML = '<b>' + data[i].LCL + '</b><i class="w3-tiny"> (Contratto:' + data[i].CN + ', ' + typeLCL + ')</i><span onclick="changeCN(this.parentElement)" class="w3-button w3-transparent w3-display-right">&times;</span>';
                document.querySelector("#addListCN").appendChild(element);
            }
        }

        document.querySelector("#loadFile").style.display = "none";
        document.querySelector("#selectLCL").style.display = "block";
    }
}
