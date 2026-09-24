/* shared loader for the 4 stock screens */
/* globals exposed for cross-script use */
window.LOC={'11':'Chennai','38':'Chennai','22':'Salem','07':'Salem'};
window.LOCNAME={'11':'Chennai · sheet/plate (11)','38':'Chennai · coil (38)','22':'Salem · subcontract (22)','07':'Salem (07)','39':'In transit (39)','50':'Ready to steel (50)','51':'Pending (51)','52':'Chennai subcontract (52)'};
window.SHOWN=['11','38','22','07','39','50','51','52'];
window.brandOf=function brandOf(name){const n=(name||'').toUpperCase();for(const b of['NMDC','JSW','SSP','AMNS','TATA','NSL','SAIL'])if(n.includes(b))return b;return 'OTHER';}
window.loadStock=async function loadStock(){
  const b='?t='+Date.now();
  const [s,c,p,me]=await Promise.all([
    fetch('mis_stock.csv'+b).then(r=>r.ok?r.text():''),
    fetch('coils.csv'+b).then(r=>r.ok?r.text():''),
    fetch('plates.csv'+b).then(r=>r.ok?r.text():''),
    fetch(BASE+'/stock/mis_meta.json'+b).then(r=>r.ok?r.json():{}).catch(()=>({}))
  ]);
  const STOCK=parseCSV(s).filter(r=>SHOWN.includes(r.WhsCode)).map(r=>({
    ItemCode:r.ItemCode,ItemName:r.ItemName,Group:r.ItemGroup,WhsCode:r.WhsCode,
    Brand:brandOf(r.ItemName),Grade:(r.Grade||'').trim(),
    T:num(r.Thickness),W:num(r.Width),L:num(r.Length),
    OnHand:num(r.OnHand),Committed:num(r.Committed),OnOrder:num(r.OnOrder)
  }));
  let clamped=0;
  STOCK.forEach(r=>{const cap=Math.max(500,r.OnHand*5);if(r.OnOrder>cap){r.OnOrder=0;clamped++;}r.Available=r.OnHand-r.Committed;});
  const COILS=parseCSV(c).map(r=>({...r,Quantity:num(r.Quantity),T:num(r.Thickness),W:num(r.Width),
    age:r.GRPODate?Math.max(0,Math.round((Date.now()-new Date(r.GRPODate))/864e5)):null}));
  const PLATES=parseCSV(p).map(r=>({...r,T:num(r.Thickness),W:num(r.Width),L:num(r.Length),TotalQty:num(r.TotalQty),TotalNos:num(r.TotalNos),Brand:brandOf(r.ItemName),Grade:(r.Grade||'').trim()}));
  setStamp(me.generated);
  return {STOCK,COILS,PLATES,META:me,clamped};
}
