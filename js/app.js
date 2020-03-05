
window.onload = function () {

var CONFIG = {
    // sourceCellLocation: 'D1',
    // targetCellLocation: 'E1',
    // worksheetName: 'Sheet1',
    sourceCellLocation: 'K6',
    targetCellLocation: 'L6',
    worksheetName: 'Content Matrix 5.0',
    highlightColor: '#4cbb17',// '#008000',//'#ffa500',
    highlightEmptyColor: '#ed2939',//'#fc1c03',
    logs: true,
    version: 'exceljs',
}
var decode_cell = XLSX.utils.decode_cell;
var encode_cell = XLSX.utils.encode_cell;

// start-up
setupConfigs();
setupDropzone();

/* MAIN FUNCTION TO START READING FILES FROM DROZONE */
//https://github.com/protobi/js-xlsx
async function readFiles(files) {

    toggleLoader();

    CONFIG.version === 'xlsx' ?
            await xlsxVersion(files) :
            await exceljsVersion(files);

    toggleLoader();
}

async function xlsxVersion(files) {
    files = Array.prototype.slice.call(files);

    // get the relevant content from each file
    let fileContents = await Promise.all( files.map(_extractData) );
    CONFIG.logs && console.log('file contents', fileContents);

    // merge that into a single store
    let store = mergeData(fileContents);
    CONFIG.logs && console.log('store', store);

    // update files based on combined store
    let updatedWorkbooks = _updateWorkbooks(fileContents, store);
    CONFIG.logs && console.log('updated workbooks', updatedWorkbooks);

    // // save as excel file
    _saveToXlsx(updatedWorkbooks);
}

async function exceljsVersion(files) {
    files = Array.prototype.slice.call(files);

    // get the relevant content from each file
    let fileContents = await Promise.all( files.map(extractData) );
    CONFIG.logs && console.log('file contents', fileContents);

    // merge that into a single store
    let store = mergeData(fileContents);
    CONFIG.logs && console.log('store', store);

    // update files based on combined store
    let updatedWorkbooks = updateWorkbooks(fileContents, store);
    CONFIG.logs && console.log('updated workbooks', updatedWorkbooks);

    // // save as excel file
    saveToXlsx(updatedWorkbooks);
}

// exceljs
function extractData(file) {
    return new Promise( (res, rej) => {
        // read the excel files
        let reader = new FileReader();
        reader.onload = function(e) {
            var arrayBuffer = reader.result;
            let workbook = new ExcelJS.Workbook();
            workbook.xlsx.load(arrayBuffer).then(function(workbook) {
                let worksheet = workbook.getWorksheet( CONFIG.worksheetName );
                let sourceColumn = worksheet.getColumn( CONFIG.sourceCellLocation.split('')[0] );
                let targetColumnName = CONFIG.targetCellLocation.split('')[0];

                let map = {
                    file: file.name,
                    workbook,
                    data: {}
                };

                sourceColumn.eachCell( (cell, rowNumber) => {
                    let target = worksheet.getCell(`${targetColumnName}${rowNumber}`);

                    if (target.text) {
                        map.data[cell.text] = target.text;
                    }
                });
                res(map)
            });
        };
        reader.onerror = function (error) {
            console.log(error);
            rej(error);
        }
        reader.readAsArrayBuffer(file);
    });
}	

// xlsx
function _extractData(file) {
    return new Promise( (res, rej) => {
        // read the excel files
        let reader = new FileReader();
        reader.onload = function(e) {
            let data = new Uint8Array(e.target.result);
            let workbook = XLSX.read(data, {type: 'array', cellStyles: true});
            let sheet = workbook.Sheets[ CONFIG.worksheetName ];
            let map = {
                file: file.name,
                workbook,
                data: {}
            };

            let sourceCoords = decode_cell(CONFIG.sourceCellLocation); // {c: 0, r: 0} A = c, 1 = r
            let targetCoords = decode_cell(CONFIG.targetCellLocation); // {c: 1, r: 0} B = c, 1 = r

            let allSourceColumnCells = getAllSourceCells(sheet, sourceCoords);
            let sourceCell, sourceCellCoords, source, target;
            while( allSourceColumnCells.length ) {
                sourceCell = allSourceColumnCells.pop();
                sourceCellCoords = decode_cell(sourceCell);

                source = sheet[ sourceCell ].v;
                target = sheet[ encode_cell( {c: targetCoords.c, r: sourceCellCoords.r } ) ];

                map.data[source] = target && target.v ? target.v : null;
            }
            res(map);
        };
        reader.onerror = function (error) {
            console.log(error);
            rej(error);
        }
        reader.readAsArrayBuffer(file);
    });
}			

// both
function mergeData(raw) {
    return raw.reduce( (final, {file, data}) => {
        let entries = Object.entries( data );
        entries.forEach( ([source, target]) => {
            if (target)
                final[source] = { target, file }
        })
        
        return final;
    }, {});
}

// exceljs
function updateWorkbooks(workbooks, store) {
    return workbooks.map( wb => {
        let workbook = wb.workbook;
        let worksheet = workbook.getWorksheet( CONFIG.worksheetName );;
        let sourceColumn = worksheet.getColumn( CONFIG.sourceCellLocation.split('')[0] );
        let targetColumnName = CONFIG.targetCellLocation.split('')[0];

        let color = '00' + (CONFIG.highlightColor.replace('#', '').toUpperCase());
        let emptyColor = '00' + (CONFIG.highlightEmptyColor.replace('#', '').toUpperCase());
        sourceColumn.eachCell( (cell, rowNumber) => {
            let source = cell.text;
            let target = worksheet.getCell(`${targetColumnName}${rowNumber}`);

            // if target already has value, skip update
            if ( !target.text && store[source] ) {
                target.value = store[source].target;

                target.style = Object.create(target.style);
                
                target.fill = {
                    type: 'pattern',
                    pattern: 'solid',
                    fgColor: {argb: color },
                    // bgColor: {argb: color },
                }
            } else if (source && !target.text) {
                target.style = Object.create(target.style);
                
                target.fill = {
                    type: 'pattern',
                    pattern: 'solid',
                    fgColor: {argb: emptyColor },
                    // bgColor: {argb: emptyColor },
                }
            }
        });
		
        return wb;
    });
}

// xlsx
function _updateWorkbooks(workbooks, store) {
    return workbooks.map( workbook => {
        let wb = workbook.workbook;
        let sheet = wb.Sheets[ CONFIG.worksheetName ];

        let sourceCoords = decode_cell(CONFIG.sourceCellLocation); // {c: 0, r: 0} A = c, 1 = r
        let targetCoords = decode_cell(CONFIG.targetCellLocation); // {c: 1, r: 0} B = c, 1 = r

        let allSourceColumnCells = getAllSourceCells(sheet, sourceCoords);
        let sourceCell, sourceCellCoords, source, target, targetCell;
        while( allSourceColumnCells.length ) {
            sourceCell = allSourceColumnCells.pop();
            sourceCellCoords = decode_cell(sourceCell);
            targetCell = encode_cell( {c: targetCoords.c, r: sourceCellCoords.r } );

            source = sheet[ sourceCell ].v;
            target = sheet[ targetCell ];

            // if target already has value, skip update
            if ( (!target || !target.v) && store[source] ) {
                let cell = { v: store[source].target, t: 's' };

                // styles
                cell.s = {
                    fill: {
                        patternType: 'solid',
                        fgColor: {rgb: CONFIG.highlightColor.replace('#', '').toUpperCase() }
                    }
                }

                sheet[ targetCell ] = cell;
            }
        }
		
        return workbook;
    });
}

// exceljs
function saveToXlsx(files) {
    files.forEach( file => {
        const wbout = file.workbook.xlsx.writeBuffer({ base64: true });
        wbout.then( buffer => {
            saveAs(
                new Blob([buffer], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" }),
                `merged_${file.file}`
            );
        });
    });
}

// xlsx
function _saveToXlsx(files) {
    //export and save file
    const wopts = { bookType:'xlsx', bookSST:true, type:'binary', cellStyles: true };

    function s2ab(s) {
        const buf = new ArrayBuffer(s.length);
        const view = new Uint8Array(buf);
        for (let i=0; i!=s.length; ++i) view[i] = s.charCodeAt(i) & 0xFF;
        return buf;
    }

    files.forEach( file => {
        const wbout = XLSX.write(file.workbook, wopts);
        saveAs(new Blob([s2ab(wbout)], {type:""}), `merged_${file.file}`);
    });
}

function getAllSourceCells(sheet, sourceCoords) {
    return Object.keys( sheet )
        .filter( cellCode => {
            if (cellCode.startsWith('!') || decode_cell(cellCode).c !== sourceCoords.c) 
                return false;
            
            return true;
        });
}

function setupConfigs() {
    let configs = Object.entries(CONFIG);

    configs.forEach( ([id, value]) => {
        let input = document.getElementById(id);
        input && ( input.value = value );
    })

    document.getElementsByClassName('controls')[0]
            .addEventListener( 'change', e => {
                CONFIG[ e.target.name ] = e.target.value;
            });
}

/* DROPZONE */
function setupDropzone() {
    document.getElementById('excel').addEventListener('change', onInputChange, true);
    document.getElementById('dropzone').addEventListener('dragover', onDragOver, true);
    document.getElementById('dropzone').addEventListener('drop', onInputChange, true);
}

function onDragOver(e) {
    e.stopPropagation();
    e.preventDefault();
    e.dataTransfer.dropEffect = 'copy';
  }

function onInputChange(e) {
    e.stopPropagation();
    e.preventDefault();
    
    //get the files
    var files;
    if(e.dataTransfer) {
    files = e.dataTransfer.files;
    
    var fileName = '';
        if( files && files.length > 1 )
            {fileName = ( document.getElementById('excel').getAttribute( 'data-multiple-caption' ) || '' ).replace( '{count}', files.length );}
        else
            fileName = e.dataTransfer.files[0].name;

        if( fileName )
            document.querySelector('#dropzone label span').innerHTML = fileName;
        
        document.getElementById('excel').classList.add("input--filled");
    } else if(e.target) {
        files = e.target.files;
    }

    readFiles(files);    
}

/* loader */
function toggleLoader() {
    document.getElementById('loader').classList.toggle('on');
}

}
