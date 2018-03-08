var rABS = true; // true: readAsBinaryString ; false: readAsArrayBuffer
function handleFile(e) {
  var files = e.target.files, f = files[0];
  var fileName = files[0].name;
  var outputFileName = fileName.replace('xlsx', 'xml');
  var reader = new FileReader();
  reader.onload = function(e) {
    var data = e.target.result;
    if(!rABS) data = new Uint8Array(data);

    var workbook = XLSX.read(data, {type: rABS ? 'binary' : 'array'});
    var worksheet = workbook.Sheets[workbook.SheetNames[0]];
    var data = XLSX.utils.sheet_to_json(worksheet);
    var outputData = {
      _declaration: {
        _attributes: {
          version: '1.0',
          encoding: 'UTF-8'
        }
      },
      program: {
        instituce: []
      }
    };
    var instituce = [];
    var cleneni;

    data.forEach(function(row) {
      // radek instituce
      if (row['instituce-nazev']) {
        // zacina nova instituce, vlozim predchozi
        if (instituce.hlavicka) {
          outputData.program.instituce.push(instituce);
        }
        instituce = {
          hlavicka: {
            nazev: {
              _text: row['instituce-nazev']
            },
            kontakt: {
              _text: row['instituce-kontakty']
            },
            poznamka: {
              _text: row['instituce-poznamky-nepovinne']
            }
          },
          udalosti: {
            udalost: []
          }
        };
      }

      // radek cleneni
      if (row['podrobnejsi-cleneni-nepovinne']) {
        cleneni = row['podrobnejsi-cleneni-nepovinne'];
      }

      // radek udalosti
      if (row['datum']) {
        var udalost = {
          cleneni: cleneni,
          datum: row['datum'],
          cas: row['cas'],
          'udalost-nazev': row['nazev'],
          anotace: row['anotace']
        };
        if (cleneni) {
          cleneni = ''
        }
        else {
          delete udalost.cleneni;
        }
        instituce.udalosti.udalost.push(udalost)
      }
    });
    outputData.program.instituce.push(instituce);
    var output = js2xml(outputData, {
      spaces: '\t',
      compact: true
    });
    download(outputFileName, output);
  };
  if(rABS) reader.readAsBinaryString(f); else reader.readAsArrayBuffer(f);
}
document.getElementById('fileInput').addEventListener('change', handleFile, false);

function download(filename, text) {
  var element = document.createElement('a');
  element.setAttribute('href', 'data:text/plain;charset=utf-8,' + encodeURIComponent(text));
  element.setAttribute('download', filename);

  element.style.display = 'none';
  document.body.appendChild(element);

  element.click();

  document.body.removeChild(element);
}
