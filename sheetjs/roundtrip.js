const X=require('xlsx'), fs=require('fs');
const buf=fs.readFileSync('rt_orig.xlsx');
for (const [name,ropt,wopt] of [
  ["default",{type:'buffer'},{bookType:'xlsx',type:'buffer'}],
  ["cellStyles+cellDates",{type:'buffer',cellStyles:true,cellDates:true},{bookType:'xlsx',type:'buffer',cellStyles:true}],
]) {
  const wb=X.read(buf,ropt);
  const ws=wb.Sheets['Data'];
  console.log(`\n== ${name}`);
  console.log("  D2:",JSON.stringify(ws['D2']),"  E2:",JSON.stringify(ws['E2']));
  console.log("  C6:",JSON.stringify(ws['C6']),"  C7:",JSON.stringify(ws['C7']));
  console.log("  merges:",JSON.stringify(ws['!merges']));
  fs.writeFileSync(`rt_${name.replace(/\W/g,'')}.xlsx`, X.write(wb,wopt));
}
