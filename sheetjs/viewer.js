// What does a viewer actually put on screen for each stage of the pipeline?
const X=require('xlsx'), fs=require('fs');
const cells=['E2','D2','D3'];
for (const f of ['rt_orig.xlsx','rt_default.xlsx','rt_cellStylescellDates.xlsx']) {
  const wb=X.read(fs.readFileSync(f),{type:'buffer'});     // a viewer's default read
  const ws=wb.Sheets['Data'];
  console.log(f.padEnd(30), cells.map(c=>`${c}: v=${ws[c].v} w=${JSON.stringify(ws[c].w)}`).join('  '));
}
