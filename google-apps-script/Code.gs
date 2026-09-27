/**
 * Project 39 — Google Sheets + Apps Script
 * Flow: Member -> Gate01 PASS -> Journey -> Product -> Market -> Sales -> Evidence
 */
const CFG={S:{MEMBERS:'Members',GATE:'Gate01',JOURNEY:'Journey',PRODUCTS:'Products',SHOPS:'Shops',SALES:'Sales',EVIDENCE:'Evidence',AUDIT:'AuditLog'}};

function doGet(){return HtmlService.createHtmlOutputFromFile('Index').setTitle('บ้านแห่งรอยยิ้ม | Project 39').setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);}
function setupProject39(){
 const ss=SpreadsheetApp.getActive();
 const schemas={
  Members:['MemberID','FullName','Village','Occupation','GateStatus','GateVerifiedAt','GateMissingInfo','GateInspector','GateEvidenceID','CreatedAt','UpdatedAt'],
  Gate01:['AuditID','MemberID','OldStatus','NewStatus','OldMissingInfo','NewMissingInfo','OldEvidenceID','NewEvidenceID','Inspector','VerifiedAt','CreatedAt'],
  Journey:['JourneyID','MemberID','Occupation','ProductID','ShopID','Status','CreatedAt','UpdatedAt'],
  Products:['ProductID','ProductName','Brand','Category','Unit','Status','UpdatedAt'],
  Shops:['ShopID','ShopName','Village','Contact','Status','UpdatedAt'],
  Sales:['SaleID','MemberID','ProductID','ShopID','SaleDate','Qty','UnitPrice','Revenue','PaymentStatus','EvidenceID','CreatedAt'],
  Evidence:['EvidenceID','EvidenceType','RefID','FileURL','Note','UploadedBy','CreatedAt'],
  AuditLog:['LogID','Entity','EntityID','Action','Actor','Detail','CreatedAt']
 };
 Object.keys(schemas).forEach(n=>{let sh=ss.getSheetByName(n)||ss.insertSheet(n);if(sh.getLastRow()===0)sh.appendRow(schemas[n]);sh.setFrozenRows(1);sh.getRange(1,1,1,schemas[n].length).setFontWeight('bold');});
 seedProducts_(); seedMembers_(); return 'Project 39 พร้อมใช้งาน';
}
function seedProducts_(){const sh=SpreadsheetApp.getActive().getSheetByName(CFG.S.PRODUCTS);if(sh.getLastRow()>1)return;sh.getRange(2,1,3,7).setValues([
 ['P001','กล้วยฉาบแม่มณีทอง','MAE MANEE THONG','อาหารแปรรูป','ถุง','Pilot / ต้องยืนยันมาตรฐาน',new Date()],
 ['P002','ไข่ผำสดและแปรรูป','FARMER POWER PRODUCT','อาหาร/เกษตรแปรรูป','หน่วย','Pilot / รอผลทดสอบและยืนยันราคา',new Date()],
 ['P003','บาล์มพิมเสน 15 กรัม','WISE PLUS GROW','สมุนไพร','กระปุก','รอสินค้า/ยืนยันขั้นสุดท้าย',new Date()] ]);}
function seedMembers_(){
 const sh=SpreadsheetApp.getActive().getSheetByName(CFG.S.MEMBERS);if(sh.getLastRow()>1)return;
 const data=[["M001","นายกชภัท วงศ์จันทา","ดอนเมย"],["M002","นายทรงชัย วงศ์จันทา","ดอนเมย"],["M003","นายวิเศษ วงษาภิลา","ดอนเมย"],["M004","นายวัยวุฒิ ทานะสิทธิ์","ดอนเมย"],["M005","นายมงคลสวัสดิ์ ทาตาสุข","ดอนเมย"],["M006","นายจิระวัฒน์ ตราษี","ดอนเมย"],["M007","นางยุวพงษ์ จารุภาค","ดอนเมย"],["M008","นายพรหมพิริยะ วงศ์จันทา","ดอนเมย"],["M009","นางสมทรง วงศ์จันทา","ดอนเมย"],["M010","นายศุภกฤต พินิจมนตรี","ดอนเมย"],["M011","นายสมชัย ศรีมันตะ","ดอนเมย"],["M012","นางสี พันเจริญ","ดอนเมย"],["M013","นางทองดี อรบุตร","ดอนเมย"],["M014","นางไสว อรบุตร","ดอนเมย"],["M015","นางประไม วงศ์จันทา","ดอนเมย"],["M016","นางสาวอัฉราพร วงศ์จันทา","ดอนเมย"],["M017","นางสุทิน เห็มจันทร์","น้ำปลีก"],["M018","นางจันทา นีระพันธ์","ดอนเมย"],["M019","นายปัญญา วงษาภิลา","ดอนเมย"],["M020","นางสาววนิดา วงษาภิลา","ดอนเมย"],["M021","นายวิทยา วงษาภิลา","ดอนเมย"],["M022","นางสาวรุ่งทิวา พรมภักดี","ดอนเมย"],["M023","นางสาวกัณฐิกา พรมภักดี","ดอนเมย"],["M024","นายกฤษณา พรมภักดี","ดอนเมย"],["M025","นายคาวี เห็มจันทร์","ดอนเมย"],["M026","นางสาวกิริตา กาแก้ว","ดอนเมย"],["M027","นางบุญมี สุภาพันธ์","ดอนเมย"],["M028","นายใส ทานะสิทธิ์","ดอนเมย"],["M029","นางหนูสินธ์ ทานะสิทธิ์","ดอนเมย"],["M030","นางสาวลัขณา ทานะสิทธิ์","ดอนเมย"],["M031","นางสาวสุนิสา จำปาหอม","ดอนเมย"],["M032","นางนายสันญา กองสิน","ดอนเมย"],["M033","นางนายกิจก้อง กองสิน","ดอนเมย"],["M034","นายเล็ก วงศ์จันทา","ดอนเมย"],["M035","นางสาวเพ็ญ เพ็งพา","ดอนเมย"],["M036","นายไทยรัตน์ เสรีวงศ์","ดอนเมย"],["M037","นางสาว พิสมัย ผิวทอง","ดอนเมย"],["M038","นางสาวนันทนา เพ็งพา","ดอนเมย"],["M039","นางสาวอรุณี สิงห์ตะโคตร","คำพระ"],["M040","นายประดิษฐ์ สอดส่อง","คำพระ"],["M041","นางนิตยา สอดส่อง","คำพระ"],["M042","นางรุ่งนภา ดวงใจ","ดอนเมย"],["M043","นายไพรย์ โมทอง","ดอนเมย"],["M044","นายนิรันดร์ แก้วโคตร","ดอนเมย"],["M045","นายเหลี่ยม ทองเสริญ","โนนโพธิ์"],["M046","นางแสงสว่าง ทองเสริญ","โนนโพธิ์"],["M047","นางปราณี พุทธคาม","โนนโพธิ์"],["M048","นายนิยม ตราษี","ดอนเมย"],["M049","นางศิรินาถ วงศ์จันทา","ดอนเมย"],["M050","นางโสภา วงศ์จันทา","ดอนเมย"],["M051","นายสมบูรณ์ จารุภาค","ดอนเมย"],["M052","นายอุดม สุวะมาตย์","นายม"],["M053","นางมณีทอง สุวะมาตย์","นายม"],["M054","นางสาวมีชัย มนต์แข็ง","นายม"],["M055","นางเปรมฤทัย นันทะใส","น้ำปลีก"],["M056","นางฮู้ จำปาหอม","ดอนเมย"],["M057","นางถนอม สุวไกร","ดอนเมย"],["M058","นางสาว อรัญญา ศรีสุพรรณ์","ดอนเมย"],["M059","นายสุเมธ ทาตาสุข","ดอนเมย"],["M060","นางสมหวัง บุญสอาด","นายม"]];
 const now=new Date();sh.getRange(2,1,data.length,11).setValues(data.map(x=>[...x,'','ยังไม่ตรวจ','','','','',now,now]));
}
function read_(name){const sh=SpreadsheetApp.getActive().getSheetByName(name);if(!sh||sh.getLastRow()<2)return[];const v=sh.getDataRange().getValues(),h=v.shift();return v.map(r=>h.reduce((o,k,i)=>(o[k]=r[i],o),{}));}
function gatePassed_(m){return !!(m&&m.GateStatus==='ตรวจแล้ว'&&String(m.Village).trim()&&String(m.Occupation).trim()&&String(m.GateInspector).trim()&&m.GateVerifiedAt&&String(m.GateEvidenceID).trim()&&!String(m.GateMissingInfo||'').trim());}
function getAppData(){const m=read_(CFG.S.MEMBERS),s=read_(CFG.S.SALES);return{members:m,products:read_(CFG.S.PRODUCTS),shops:read_(CFG.S.SHOPS),journey:read_(CFG.S.JOURNEY),sales:s,evidence:read_(CFG.S.EVIDENCE),dashboard:{members:m.length,passed:m.filter(gatePassed_).length,revenue:s.reduce((a,x)=>a+(Number(x.Revenue)||0),0)}};}
function saveGate01(p){
 if(p.GateStatus==='ตรวจแล้ว' && (!p.Village||!p.Occupation||!p.GateInspector||!p.GateVerifiedAt||!p.GateEvidenceID||String(p.GateMissingInfo||'').trim()))throw Error('Gate 01 PASS ต้องครบ 7 เงื่อนไข');
 const sh=SpreadsheetApp.getActive().getSheetByName(CFG.S.MEMBERS),v=sh.getDataRange().getValues(),h=v[0],i=h.reduce((o,k,n)=>(o[k]=n,o),{}),r=v.findIndex(x=>String(x[i.MemberID])===String(p.MemberID));if(r<1)throw Error('ไม่พบสมาชิก');
 const old=v[r],nr=[...old];['Village','Occupation','GateStatus','GateVerifiedAt','GateMissingInfo','GateInspector','GateEvidenceID'].forEach(k=>nr[i[k]]=p[k]||'');nr[i.UpdatedAt]=new Date();sh.getRange(r+1,1,1,h.length).setValues([nr]);
 SpreadsheetApp.getActive().getSheetByName(CFG.S.GATE).appendRow([Utilities.getUuid(),p.MemberID,old[i.GateStatus],p.GateStatus,old[i.GateMissingInfo],p.GateMissingInfo||'',old[i.GateEvidenceID],p.GateEvidenceID||'',p.GateInspector||'',p.GateVerifiedAt||'',new Date()]);
 return getAppData();
}
function createJourney(p){const m=read_(CFG.S.MEMBERS).find(x=>x.MemberID===p.MemberID);if(!gatePassed_(m))throw Error('สมาชิกยังไม่ผ่าน Gate 01');if(!p.Occupation||!p.ProductID)throw Error('ต้องมีอาชีพและ Product');SpreadsheetApp.getActive().getSheetByName(CFG.S.JOURNEY).appendRow([Utilities.getUuid(),p.MemberID,p.Occupation,p.ProductID,p.ShopID||'','ACTIVE',new Date(),new Date()]);return getAppData();}
function createSale(p){const m=read_(CFG.S.MEMBERS).find(x=>x.MemberID===p.MemberID);if(!gatePassed_(m))throw Error('สมาชิกยังไม่ผ่าน Gate 01');if(!p.ProductID||!p.SaleDate||!p.Qty||!p.UnitPrice||!p.EvidenceID)throw Error('รายรับต้องมี Product + วันที่ + จำนวน + ราคา + Evidence');SpreadsheetApp.getActive().getSheetByName(CFG.S.SALES).appendRow([Utilities.getUuid(),p.MemberID,p.ProductID,p.ShopID||'',p.SaleDate,Number(p.Qty),Number(p.UnitPrice),Number(p.Qty)*Number(p.UnitPrice),p.PaymentStatus||'รอตรวจ',p.EvidenceID,new Date()]);return getAppData();}
