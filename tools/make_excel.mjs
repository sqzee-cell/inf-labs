import fs from 'node:fs/promises';
import path from 'node:path';
import {Workbook,SpreadsheetFile} from '@oai/artifact-tool';
const root='F:/study/inf-labs',tmp=path.join(root,'tmp/excel');
await fs.mkdir(tmp,{recursive:true});
const wb=Workbook.create();
const names=['Календарь','Таблица','Функция','Поверхность','Матрицы'];
const sheets=names.map(n=>wb.worksheets.add(n));
const put=(s,a,v)=>s.getRange(a).values=v;
const formula=(s,a,v)=>s.getRange(a).formulas=v;
const head=(s,a,t)=>{s.mergeCells(a);put(s,a.split(':')[0],[[t]]);s.getRange(a).format={fill:'#E8EFF6',font:{name:'Arial',size:16,bold:true,color:'#173654'},rowHeight:32};};
for(const s of sheets){s.showGridLines=false;s.getRange('A1:Q55').format.font={name:'Arial',size:11,color:'#202020'};s.getRange('A1:Q55').format.columnWidth=14;s.getRange('A1:Q55').format.rowHeight=22;}
const cal=sheets[0];head(cal,'B2:H3','НОЯБРЬ 2026');
put(cal,'B4',[["Дубограй Артем Юрьевич · группа 615 · вариант 6"]]);
cal.mergeCells('B4:H4');
put(cal,'B6:H6',[['Пн','Вт','Ср','Чт','Пт','Сб','Вс']]);
cal.getRange('B6:H6').format={fill:'#173654',font:{color:'#FFFFFF',bold:true},horizontalAlignment:'center'};
const grid=Array.from({length:6},()=>Array(7).fill(null));
for(let d=1;d<=30;d++){let n=6+d-1;grid[Math.floor(n/7)][n%7]=d;}
put(cal,'B7:H12',grid);cal.getRange('B7:H12').format={horizontalAlignment:'center',verticalAlignment:'center',rowHeight:52,font:{size:18},borders:{preset:'all',style:'thin',color:'#D9D9D9'}};
cal.getRange('G7:H12').format.fill='#F2F2F2';cal.getRange('D8').format={fill:'#FBE2DD',font:{color:'#B02E22',bold:true}};
for(const a of ['C9','G10'])cal.getRange(a).format.fill='#FFF1CF';
cal.mergeCells('B15:H15');put(cal,'B15',[['4 ноября — День народного единства']]);
cal.mergeCells('B16:H16');put(cal,'B16',[['10 ноября — День сотрудника органов внутренних дел']]);
cal.mergeCells('B17:H17');put(cal,'B17',[['21 ноября — День работника налоговых органов']]);
cal.mergeCells('B19:H20');put(cal,'B19',[['Красный — государственный праздник; серый — суббота/воскресенье; жёлтый — профессиональный праздник.']]);cal.getRange('B19:H20').format.wrapText=true;
cal.mergeCells('B22:H22');put(cal,'B22',[['Источники дат: Минтруд России, МВД России и ФНС России']]);
cal.mergeCells('B23:H23');put(cal,'B23',[['https://mintrud.gov.ru/social/social/1366']]);
cal.mergeCells('B24:H24');put(cal,'B24',[['https://mvdmedia.ru/news/official/pozdravlenie-prezidenta-rossiyskoy-federatsii-vladimira-vladimirovicha-putina-s-dnem-sotrudnika-orga/']]);
cal.mergeCells('B25:H25');put(cal,'B25',[['https://www.nalog.gov.ru/rn05/news/activities_fts/16583225/']]);
cal.getRange('B23:H25').format.font.size=8;

const pay=sheets[1];head(pay,'A1:G2','Заработная плата — учебный пример');
put(pay,'A4:G4',[['№','Фамилия','Оклад','Премия','Сверхурочные','Налог','Начисление']]);
const surnames=['Воланд','Мелехов','Микша','Облонский','Онегин','Печорин','Преображенский','Раскольников','Ростова','Старцев'];
const salaries=[3000,2000,1900,1200,1800,2500,2900,2800,1300,1500];
put(pay,'A5:C14',surnames.map((n,i)=>[i+1,n,salaries[i]]));
put(pay,'I4:J7',[['Параметр','Значение'],['Ставка премии',.11],['Сверхурочные',.555],['Ставка налога',.13]]);
put(pay,'I9:J9',[['Учебный порог',2000]]);pay.getRange('J5:J7').setNumberFormat('0.0%');
pay.mergeCells('I11:M13');put(pay,'I11',[['Ставки и оклады — данные учебного образца, а не действующие нормативы оплаты труда.']]);pay.getRange('I11:M13').format.wrapText=true;
formula(pay,'D5:G5',[['=ROUND(C5*$J$5,2)','=ROUND(C5*$J$6,2)','=ROUND(SUM(C5:E5)*$J$7,2)','=SUM(C5:E5)-F5']]);pay.getRange('D5:G14').fillDown();
put(pay,'B15:B20',[['ИТОГО'],['Средний показатель'],['MAX'],['MIN'],['Больше порога'],['Меньше порога']]);
for(const c of ['C','D','E','F','G'])formula(pay,`${c}15:${c}18`,[[`=SUM(${c}5:${c}14)`],[`=AVERAGE(${c}5:${c}14)`],[`=MAX(${c}5:${c}14)`],[`=MIN(${c}5:${c}14)`]]);
formula(pay,'C19:C20',[['=COUNTIF(C5:C14,">"&$J$9)'],['=COUNTIF(C5:C14,"<"&$J$9)']]);
pay.getRange('C5:G18').setNumberFormat('#,##0.00" р."');pay.getRange('C19:C20').setNumberFormat('0');
pay.getRange('A4:G20').format.borders={preset:'all',style:'dashDot',color:'#000080'};
pay.getRange('A4:G20').format.borders={preset:'outside',style:'mediumDashDot',color:'#333F50'};
const fills=['#CCFFFF','#CCFFFF','#CCFFFF','#CCFFCC','#CCFFCC','#CCFFCC','#CCFFFF'];
for(let i=0;i<7;i++)pay.getRange(`${String.fromCharCode(65+i)}5:${String.fromCharCode(65+i)}14`).format.fill=fills[i];
pay.getRange('A4:G4').format={fill:'#FFFFCC',font:{bold:true},wrapText:true,rowHeight:48};pay.getRange('A15:G15').format={fill:'#FFFFCC',font:{bold:true}};
pay.getRange('A5:B14').format.horizontalAlignment='left';pay.getRange('A5:B14').format.verticalAlignment='center';
pay.getRange('C5:E14').format.horizontalAlignment='right';pay.getRange('C5:E14').format.verticalAlignment='center';
pay.getRange('F5:G14').format.horizontalAlignment='left';pay.getRange('F5:G14').format.verticalAlignment='top';
pay.getRange('A1:A20').format.columnWidth=6;pay.getRange('B1:B20').format.columnWidth=26;pay.getRange('C1:G20').format.columnWidth=18;pay.getRange('I1:I20').format.columnWidth=22;

const fun=sheets[2];head(fun,'A1:I2','Табулирование функции одной переменной');
put(fun,'A4:F5',[['a','b','c','x₀, градусы','xₙ, градусы','h, градусы'],[2,3,2,10,360,10]]);
fun.mergeCells('A7:K7');put(fun,'A7',[['y = a·sin(b·x) / ∛(cos(x)+c) + log₃(x); углы тригонометрических функций переведены в радианы.']]);
put(fun,'A9:B9',[['x, градусы','y']]);formula(fun,'A10',[['=$D$5']]);formula(fun,'A11',[['=A10+$F$5']]);fun.getRange('A11:A45').fillDown();
formula(fun,'B10',[['=$A$5*SIN(RADIANS($B$5*A10))/(COS(RADIANS(A10))+$C$5)^(1/3)+LOG(A10,3)']]);fun.getRange('B10:B45').fillDown();fun.getRange('B10:B45').setNumberFormat('0.000000');
const chart=fun.charts.add('scatter',fun.getRange('A9:B45'));chart.title='y(x), x в градусах';chart.hasLegend=false;chart.setPosition('D10','N28');chart.xAxis.title.text='x, градусы';chart.yAxis.title.text='y';chart.series.items[0].line={fill:'#174E85',style:'solid',width:2};

const surf=sheets[3];head(surf,'A1:N2','Поверхность z = x²/a² − y²/b²');put(surf,'A4:D5',[['a, градусы','b, градусы','Шаг x','Шаг y'],[45,30,15,15]]);
put(surf,'A7',[["y / x, градусы"]]);put(surf,'B7:N7',[Array.from({length:13},(_,i)=>-90+i*15)]);put(surf,'A8:A20',Array.from({length:13},(_,i)=>[-90+i*15]));
formula(surf,'B8',[['=B$7^2/$A$5^2-$A8^2/$B$5^2']]);surf.getRange('B8:N8').fillRight();surf.getRange('B8:N20').fillDown();surf.getRange('B8:N20').setNumberFormat('0.0000');
surf.getRange('B8:N20').conditionalFormats.add('colorScale',{colors:['#C6DBEF','#FFFFFF','#F4B183'],thresholds:['min','50%','max']});
surf.getRange('A1:N28').format.columnWidth=10;surf.mergeCells('A23:N24');put(surf,'A23',[['Здесь x и y входят в формулу только в квадрате. Значения x, y, a, b заданы в одинаковых единицах (градусах); преобразования тригонометрических аргументов нет.']]);surf.getRange('A23:N24').format.wrapText=true;

const mat=sheets[4];head(mat,'A1:P2','Матричное выражение (A − B)·C²');
put(mat,'B3',[['Матрица_A']]);put(mat,'F3',[['Матрица_B']]);put(mat,'J3',[['Матрица_C']]);put(mat,'N3',[['Матрица_D']]);
put(mat,'B4:D6',[[3,1,2],[0,4,1],[2,1,5]]);put(mat,'F4:H6',[[1,0,1],[0,2,0],[1,0,2]]);put(mat,'J4:L6',[[1,2,0],[0,1,1],[1,0,2]]);put(mat,'N4:P6',[[1,0,0],[0,1,0],[0,0,1]]);
put(mat,'B9',[['A − B']]);put(mat,'F9',[['C² = C·C']]);put(mat,'J9',[['Результат']]);
for(let r=0;r<3;r++)for(let c=0;c<3;c++){
 const col=n=>String.fromCharCode(65+n);
 formula(mat,`${col(1+c)}${10+r}`,[[`=${col(1+c)}${4+r}-${col(5+c)}${4+r}`]]);
 formula(mat,`${col(5+c)}${10+r}`,[[`=J${4+r}*${col(9+c)}4+K${4+r}*${col(9+c)}5+L${4+r}*${col(9+c)}6`]]);
 formula(mat,`${col(9+c)}${10+r}`,[[`=B${10+r}*${col(5+c)}10+C${10+r}*${col(5+c)}11+D${10+r}*${col(5+c)}12`]]);
}
mat.mergeCells('B15:P16');put(mat,'B15',[['Матрицы 3×3 выбраны самостоятельно. Квадрат C² — матричное произведение C·C. Диапазон Матрица_D задан по требованию методички, но в выражении варианта 6 не участвует.']]);mat.getRange('B15:P16').format.wrapText=true;
for(const s of [fun,surf,mat]){s.getUsedRange().format.verticalAlignment='center';}
wb.recalculate();
console.log((await wb.inspect({kind:'table',range:'Таблица!B15:G20',include:'values,formulas',tableMaxRows:6,tableMaxCols:6,maxChars:1700})).ndjson);
console.log((await wb.inspect({kind:'match',searchTerm:'#REF!|#DIV/0!|#VALUE!|#NAME\\?|#NUM!|#NULL!',options:{useRegex:true,maxResults:20},summary:'Formula error scan',maxChars:1500})).ndjson);
for(const [i,s] of sheets.entries()){
 const ranges=['B2:H20','A1:J20','A1:N28','A1:N24','B3:P16'];
 try {const blob=await wb.render({sheetName:names[i],range:ranges[i],scale:1.5,format:'png'});await fs.writeFile(path.join(tmp,`sheet-${i+1}.png`),new Uint8Array(await blob.arrayBuffer()));}catch(e){console.log(`Preview ${names[i]}: ${e.message}`);}
}
const output=await SpreadsheetFile.exportXlsx(wb);await output.save(path.join(tmp,'base.xlsx'));
console.log('Saved base.xlsx');
