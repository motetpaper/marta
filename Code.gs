// Code.gs

// create add-on menu
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Download as Markdown')
    .addItem('Download as Markdown text file', 'getMartaObject')
    .addToUi();
}

// yes.
function onInstall() {
  onOpen();
}

// get the marta object (custom google sheets map)
// process the map into markdown text
function getMartaObject() {
  const marta = JSON.stringify(mkmarta());
  // working marta client-side handoff prototype
  const ui =  SpreadsheetApp.getUi();
  const html = HtmlService.createHtmlOutput(`
  <!doctype html>
  <html>
  <head></head>
  <body><script defer type="text/javascript">
  document.body.onload = () => {
    marta(${marta});
  }
  function marta(obj) {
    const sheet = obj;
    const now = Date.now();
    const outfile = sheet.name+'-'+now+'.md.txt';
    let thetable = sheet.dvals.map((a)=>{
      return a.map((str)=>{
        return '| ' + str + ' ';
      });
    });

    console.table(thetable)

    // removes the hidden rows
    thetable = thetable.map((a,i)=> {
      return a.filter((b,j)=>{
        return !sheet.hidecols[j][1];
      });
    });    

    // removes hidden rows
    thetable = thetable.filter((a,i)=>{
        return !sheet.hiderows[i][1];
    });

    const alignRow = thetable[0].map((a,i)=>{
      return getAlignRow(sheet.aligns[0][i]);
    });

    console.log(alignRow);
    // adds alignments to the markdown table
    thetable.splice(1,0,alignRow);

    
    const outdata = thetable.map((a)=>a.join('')+'|').join('\\n')+'\\n'
    const blob = new Blob([outdata], {
      type: 'text/plain'
    });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = outfile;
    a.click();

    // function ..
    

    // returns column alignment in markdown syntax
    function getAlignmentMarkdown(str) {
      switch(str.toLowerCase()) {
      case 'left':
      case 'general-left':
        return ':--'
      case 'left':
      case 'general-left':
        return ':--'
      case 'left':
      case 'general-left':
        return ':--'
      default:
        return ':--:';
      }
    }

    // returns alignment in markdown syntax
    function getAlignRow(str) {
      const align = getAlignmentMarkdown(str);
      return '| ' + align + ' ';
    }    
  }
  </script>
  </body>
  </html>
  `);
  ui.showSidebar(html);
}



function mkmarta() {

  const obj = {};

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sht = ss.getActiveSheet();
  const rng = sht.getDataRange();
  const rows = rng.getNumRows();
  const cols = rng.getNumColumns();
 
  // sheet data area
  obj.name = sht.getSheetName();
  obj.notation = rng.getA1Notation();

  obj.rows = rows;
  obj.cols = cols;


  obj.lastcol = rng.getLastColumn();
  obj.lastrow = rng.getLastRow();

  obj.fcols = sht.getFrozenColumns();
  obj.frows = sht.getFrozenRows(); 
  obj.aligns = rng.getHorizontalAlignments();

  // hidden rows

  obj.hiderows = [];
  let isRowHidden = null;
  for(let j = 1; j < rows+1; j++) {
    isRowHidden = sht.isRowHiddenByFilter(j) || sht.isRowHiddenByUser(j);
    obj.hiderows.push([`ROW-${j}`, isRowHidden]);
  }
  
  // hidden columns

  obj.hidecols = [];
  let isColHidden = null;
  for(let i = 1; i < cols+1; i++) {
    isColHidden = sht.isColumnHiddenByUser(i);
    obj.hidecols.push([`COL-${i}`,isColHidden]);
  }  

  // cell values area
  obj.dvals = rng.getDisplayValues();
  obj.vals = rng.getValues();
  
  // fonts area
  obj.fweights = rng.getFontWeights();
  obj.fstyles = rng.getFontStyles();
  obj.fontlines = rng.getFontLines();    
  obj.ffams = rng.getFontFamilies(); 
  obj.rtvals = rng.getRichTextValues();
  obj.tstyles = rng.getTextStyles();
  
  // detects if text style is underline
  obj.ts_underline = obj.tstyles.map((a,j)=>{
    return a.map((b,i)=>{
      return b.isUnderline();
    });
  });

  // detects if text style is bold
  obj.ts_bold = obj.tstyles.map((a,j)=>{
    return a.map((b,i)=>{
      return b.isBold();
    });
  });

  // detects if text style is strikethrough (line-through)
  obj.ts_strikethrough = obj.tstyles.map((a,j)=>{
    return a.map((b,i)=>{
      return b.isStrikethrough();
    });
  });


  // detects if text style is italic
  obj.ts_italic = obj.tstyles.map((a,j)=>{
    return a.map((b,i)=>{
      return b.isItalic();
    });
  });    

  debug_tostring(obj.underline);

  // validations area
  obj.valids = rng.getDataValidations();

  // detects if validation is checkbox
  obj.ischeckbox = obj.valids.map((a,j)=>{
    return a.map((b,i)=>{
      return (!!b) ? b.getCriteriaType().toJSON() === 'CHECKBOX': false;
    });
  });

  // formulas area
  obj.formulas = rng.getFormulas();
  obj.formulasrc = rng.getFormulasR1C1();

  // footnotes area
  obj.notes = rng.getNotes();
  obj.hasnotes = obj.notes.map((a,j)=>{
    return a.map((b,i)=>{
      return !!b.length;
    });
  });

  return obj;
}
