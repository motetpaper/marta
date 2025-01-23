// marta.js
// (prototype, client-side)
// job    : converts google sheets "mock export" object into markdown table
// git    : https://github.com/motetpaper/marta
// lic    : MIT

const fs = require('fs');

/**
 * Sample Tests:
 * $ node marta.js; diff b1.md b3.md
 * $ node marta.js; diff b2.md b3.md
 *
 **/
// Sample Mock Objects
// https://github.com/motetpaper/gsheets-mock-objects
// 'mock7-pretty.json'
// 'mock8-pretty.json'
fs.readFile('mock8-pretty.json', 'utf8', (err,data)=> {

  const sheet = JSON.parse(data);

  const rows = sheet.dvals.length;
  const cols = sheet.dvals[0].length;

  console.log('sheet-rows %s', rows);
  console.log('sheet-cols %s', cols);

  let thetable = sheet.dvals.map((a,j)=>{
    return a.map((str,i)=>{
//      console.log('R%sC%s: %s',(1+j),(1+i), str);

      // adds code syntax to markdown text
      if(isMonospace(sheet.ffams[j][i])) {
        str = `\`${str}\``;
      }

      // adds bold syntax to markdown text
      if(isBold(sheet.fweights[j][i])) {
        str = `**${str}**`;
      }

      // adds italic syntax to markdown text
      if(isItalic(sheet.fstyles[j][i])) {
        str = `*${str}*`;
      }

      // adds strikethrough syntax to markdown text
      if(isStrikethrough(sheet.fontlines[j][i])) {
        str = `~~${str}~~`;
      }

      /**
       * NOTE: The strikethrough syntax for
       * for markdown varies by platform.
       * Support for more platforms may require
       * more detailed documentation.
       */

      return `| ${str} `;
    });
  });

  debug_outdata('b1.md',thetable)

  const trows = thetable.length;
  const tcols = thetable[0].length;

  console.log('table-rows %s', trows);
  console.log('table-cols %s', tcols);

  // removes hidden rows
  thetable = thetable.filter((a,i)=>{
      return !sheet.hiderows[i][1];
  });

  // removes the hidden rows
  thetable = thetable.map((a,i)=> {
    return a.filter((b,j)=>{
      return !sheet.hidecols[j][1];
    });
  });


  const zrows = thetable.length;
  const zcols = thetable[0].length;

  console.log('final-rows %s', zrows);
  console.log('final-cols %s', zcols);

  debug_outdata('b2.md',thetable)

  // finishing area

  // the first row of the remaining
  // columns determines the alignments
  const alignRow = thetable[0].map((a,i)=>{
    return getAlignRow(sheet.aligns[0][i]);
  });

  // adds alignments to the markdown table
  thetable.splice(1,0,alignRow);

  debug_outdata('b3.md',thetable)
  debug_outdata(`motet-${sheet.name}.md`,thetable)
});



/**
 * helper functions
 */

// saves the table outdata as outfile
function debug_outdata(outfile, thetable) {
  const outdata = debug_tostring(thetable);
  fs.writeFile(outfile, outdata, (err)=>{if(err) throw err});
}

// converts table array to a string
function debug_tostring(thetable) {
  return thetable.map((a)=>a.join('')+'|').join('\n')+'\n';
}

// returns true if font line is stikethrough; otherwise, false.
function isStrikethrough(str) {
  return !!str && (str === 'line-through') ? true : false;
}

// returns true if font weight is bold; otherwise, false.
function isBold(str) {
  return !!str && (str === 'bold') ? true : false;
}

// returns true if font style is italic; otherwise, false.
function isItalic(str) {
  return !!str && (str === 'italic') ? true : false;
}

// uses fonts list found on fonts.google.com
// returns true if font family is monospaced; otherwise, false.
function isMonospace(str) {
  return `
roboto mono
inconsolata
source code pro
ibm plex mono
nanum gothic coding
jetbrains mono
space mono
vt323
courier prime
dm mono
ubuntu mono
pt mono
doto
geist mono
fira mono
cousine
share tech mono
fira code
anonymous pro
overpass mono
sixtyfour convergence
major mono display
cutive mono
oxygen mono
azeret mono
b612 mono
nova mono
syne mono
reddit mono
lekton
xanh mono
martian mono
fragment mono
chivo mono
monofett
red hat mono
lxgw wenkai mono tc
kode mono
ubuntu sans mono
m plus 1 code
spline sans mono
sono
sometype mono
sixtyfour
workbench
victor mono
`.trim()
 .split('\n')
 .indexOf(str.toLowerCase()) > -1;
}

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
  return `| ${align} `;
}
