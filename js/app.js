
window.onload = function () {

const CONFIG = {
    // sourceCellLocation: 'D1',
    // targetCellLocation: 'E1',
    // worksheetName: 'Sheet1',
    sourceCellLocation: 'K6',
    targetCellLocation: 'L6',
    worksheetName: 'Content Matrix 5.0',
    
    highlightColor: '#4cbb17',// '#008000',//'#ffa500',
    highlightEmptyColor: '#ed2939',//'#fc1c03',
    logs: false,
    version: 'exceljs', // xlsx value will trigger another library, but it will not work with the latest changes, such as fuzzy
    fuzzy: true,
    acceptableFuzzyRating: 0.5,
    highlightFuzzyColor: '#297eed'
};

const sideEffects = {
    _exec(e) {
        let sideEffect = e.target.dataset.sideEffect,
            args = e.target.dataset.sideEffectArgs;
        
        args && (args = args.split(','));

        if (sideEffects[ sideEffect ])
            sideEffects[ sideEffect ] ( e, args );
        else
            console.log(`Side-effect ${sideEffect} does not exist.`);
    },

    hideRelated(e, args) {
        args = args.map( arg => `label[for=${arg}]` ).join(',');

        let els = document.querySelectorAll(args);

        if (e.target.checked)
            els.forEach( el => el.classList.add('visible'));
        else
            els.forEach( el => el.classList.remove('visible'));
    }
};

var toast = {
    timer : null,
    useTimer : false,
    timerInterval: 3000,
    
    class: 'visible',

    init : function () {
        // SET DISMISS EVENT
        document.getElementsByClassName('toast-dismiss')[0].addEventListener('click', toast.hide);
    },

    show : function (msg) {
        // SET MESSAGE STRING
        let html = `
        ${msg.name ? `<h5>${msg.name}</h5>` : ''}
        <h3>I failed!</h3>
        ${msg.message ? `<p>${msg.message}</p>` : ''}
        `;
        // SET MESSAGE + SHOW BOX
        document.getElementsByClassName("toast-message")[0].innerHTML = html;
        document.getElementById("toast").classList.add( toast.class );

        // RESET TIMER IF STILL RUNNING
        if (toast.useTimer && toast.timer != null) {
            clearTimeout(toast.timer);
            toast.timer = null;
        }

        // SET DISPLAY TIME HERE
        toast.useTimer && ( toast.timer = setTimeout(toast.hide, toast.timerInterval) ); 
    },

    hide : function () {
        document.getElementById("toast").classList.remove( toast.class );
        
        if (toast.useTimer) {
            clearTimeout(toast.timer);
            toast.timer = null;
        }
    }
}

// start-up
setupConfigs();
setupDropzone();
toast.init();

/* MAIN FUNCTION TO START READING FILES FROM DROZONE */
async function readFiles(files) {

    toggleLoader();

    CONFIG.version === 'xlsx' ?
            await xlsxVersion(files) :
            await exceljsVersion(files);

    toggleLoader();
}

// exceljs 3.4
async function exceljsVersion(files) {
    try {
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
    } catch (error) {
        handleError(error);
    }
}

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
            }).catch( rej );
        };
        reader.onerror = function (error) {
            console.log(error);
            rej(error);
        }
        reader.readAsArrayBuffer(file);
    });
}

function updateWorkbooks(workbooks, store) {
    return workbooks.map( wb => {
        let workbook = wb.workbook,
            worksheet = workbook.getWorksheet( CONFIG.worksheetName ),
            sourceColumn = worksheet.getColumn( CONFIG.sourceCellLocation.split('')[0] ),
            targetColumnName = CONFIG.targetCellLocation.split('')[0],
            exactMatchColor = '00' + (CONFIG.highlightColor.replace('#', '').toUpperCase()),
            fuzzyMatchColor = '00' + (CONFIG.highlightFuzzyColor.replace('#', '').toUpperCase()),
            emptyColor = '00' + (CONFIG.highlightEmptyColor.replace('#', '').toUpperCase());
        
        sourceColumn.eachCell( (cell, rowNumber) => {
            let sourceText = cell.text,
                target = worksheet.getCell(`${targetColumnName}${rowNumber}`),
                color = null;

            // cancel if:
            // - no source text
            // - source text and target text
            if (    
                    !sourceText ||
                    (sourceText && target.text)
               ) return false;

            // 100% match
            if (store[sourceText]) {
                target.value = store[sourceText].target;
                color = exactMatchColor;
            }
            // fuzzy match
            else if ( CONFIG.fuzzy ) {
                let { bestMatch } = compare.findBestMatch( sourceText, Object.keys(store) );
                
                if (bestMatch.rating >= CONFIG.acceptableFuzzyRating) {
                    target.value = store[bestMatch.target].target;

                    target.note = {
                        texts: [
                            ...diff(bestMatch.target, sourceText),
                            {'font': {'size': 12, 'color': {'theme': 1}, 'name': 'Calibri', 'family': 2, 'scheme': 'minor'}, 'text': '\r\n\r\nFuzzy percent: \r\n'},
                            {'font': {'bold': true, 'size': 12, 'color': {'theme': 1}, 'name': 'Calibri', 'scheme': 'minor'}, 'text': `${ (bestMatch.rating * 100).toFixed(2) }%\r\n`},
                            {'font': {'size': 12, 'color': {'theme': 1}, 'name': 'Calibri', 'family': 2, 'scheme': 'minor'}, 'text': 'Matched source: \r\n'},
                            {'font': {'bold': true, 'size': 12, 'color': {'theme': 1}, 'name': 'Calibri', 'scheme': 'minor'}, 'text': `${bestMatch.target} \r\n`}
                        ],
                        shapeId: 2
                    }

                    color = fuzzyMatchColor;
                } else
                    color = emptyColor;
            }
            // if no 100% and no fuzzy
            else {
                color = emptyColor;
            }

            // fix the styles
            target.style = Object.create(target.style);
            
            target.fill = {
                type: 'pattern',
                pattern: 'solid',
                fgColor: {argb: color }
            }
        });
		
        return wb;
    });
}

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

// xlsx = https://github.com/protobi/js-xlsx
async function xlsxVersion(files) {
    try {
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
    } catch (error) {
        handleError(error);
    }
}

function _extractData(file) {
    return new Promise( (res, rej) => {
        const decode_cell = XLSX.utils.decode_cell;
        const encode_cell = XLSX.utils.encode_cell;
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

function _updateWorkbooks(workbooks, store) {
    const decode_cell = XLSX.utils.decode_cell;
    const encode_cell = XLSX.utils.encode_cell;
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
    const decode_cell = XLSX.utils.decode_cell;
    return Object.keys( sheet )
        .filter( cellCode => {
            if (cellCode.startsWith('!') || decode_cell(cellCode).c !== sourceCoords.c) 
                return false;
            
            return true;
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


/* CONFIG */
function setupConfigs() {
    // set up proper excel parser
    let parser = CONFIG.version;
    let script = document.createElement('script');
    script.src = `./js/xlsx/${parser}.${CONFIG.logs ? 'full' : 'min'}.js`;
    document.body.appendChild(script);
    
    //
    let configs = Object.entries(CONFIG);

    // setup event listener
    let controls = document.getElementsByClassName('controls')[0];
    controls.addEventListener( 'change', e => {
        let val = e.target.type === 'checkbox' ? e.target.checked : e.target.value;
        
        // side-effects
        if (e.target.dataset.sideEffect)
            sideEffects._exec(e);

        // update
        CONFIG[ e.target.name ] = val;

        CONFIG.logs && console.log('Event data', e, val);
        CONFIG.logs && console.log('CONFIG was updated', CONFIG);
    }, true);
    
    // update based on default config
    configs.forEach( ([id, value]) => {
        let input = document.getElementById(id);

        if (input) {
            if (input.type === 'checkbox') {
                input.checked = value;
            }
            else input.value = value;

            input.dispatchEvent( new Event('change') )
        }
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

/* LOADER */
function toggleLoader() {
    document.getElementById('loader').classList.toggle('on');
}

/* ERRORS */
function handleError(error) {
    let message = null;
    switch (error.message) {
        case 'Cannot read property \'getColumn\' of undefined':
            message = 'The sheet you have selected does not seem to exist in some of the files. Please double-check.';
            break;
        default:
            message = 'Something went wrong. See the browser console for details.';
            break;
    }

    toast.show( { name: error.name, message } );

    console.error(error);
}

}
