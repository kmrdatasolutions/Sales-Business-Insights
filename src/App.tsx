import * as React from "react";
import { useEffect, useMemo, useRef, useState } from "react";
import {
  BarChart3, Box, ChevronLeft, ChevronRight, CircleDollarSign, Database,
  FileSpreadsheet, FileText, Globe2, Home, Layers3, LayoutDashboard,
  Package, RefreshCw, Settings2, ShoppingCart, Sparkles, Table2, Target,
  Truck, UploadCloud, Users, X, Zap, TrendingUp, TrendingDown, AlertTriangle,
  CheckCircle2, ArrowUpRight, ArrowDownRight, Sun, Moon
} from "lucide-react";
import {
  Bar, BarChart, CartesianGrid, Cell, Line, LineChart, Pie, PieChart,
  ResponsiveContainer, Scatter, ScatterChart, Tooltip, XAxis, YAxis
} from "recharts";
import type { Dataset, Filters } from "./types";
import * as XLSX from "xlsx";
import jsPDF from "jspdf";
import PptxGenJS from "pptxgenjs";
import {
  applyFilters, avgBy, categoricalColumns, currency, findColumn,
  groupAvg, groupSum, inferMetrics, number, sumBy, trend, uniqueValues
} from "./lib/analytics";
import { makeDataset, parseFile } from "./lib/parsers";

type Page = "landing" | "overview" | "product" | "regional" | "insights";
const pageTitles: Record<Page, string> = {
  landing: "Landing",
  overview: "Executive Overview",
  product: "Product & Customer",
  regional: "Regional & Operations",
  insights: "Business Insights"
};

const demoRows = [
  ["2017-01-12","West","Consumer","Technology","Laptop Pro","Acme Corp","$18200",18,6200,2],
  ["2017-02-03","East","Corporate","Technology","Laser Printer","Beta LLC","$12400",12,4100,3],
  ["2017-03-19","West","Consumer","Furniture","Executive Chair","Acme Corp","$8600",20,2500,4],
  ["2017-04-11","Central","Home Office","Office Supplies","Binders","Delta Inc","$3200",44,900,5],
  ["2017-05-08","South","Corporate","Technology","Monitor 27","Gamma Ltd","$15600",15,4800,3],
  ["2017-06-21","West","Consumer","Technology","Laptop Pro","Acme Corp","$22100",21,7600,2],
  ["2017-07-09","East","Home Office","Furniture","Standing Desk","Omega Inc","$9700",9,3200,4],
  ["2017-08-14","Central","Corporate","Office Supplies","Storage","Delta Inc","$5400",31,1800,6],
  ["2017-09-18","West","Consumer","Technology","Monochrome Laser Printer","Beeline","$20400",28,6900,2],
  ["2017-10-04","South","Corporate","Furniture","Conference Table","Gamma Ltd","$13200",8,4400,5],
  ["2017-11-17","East","Consumer","Office Supplies","Accessories","Acme Corp","$7800",37,2600,3],
  ["2017-12-22","West","Consumer","Technology","Phone","Beeline","$11200",19,3500,2]
].map(r => ({
  "Order Date": r[0], Region: r[1], Segment: r[2], Category: r[3],
  "Product Name": r[4], "Customer Name": r[5], Sales: r[6],
  Quantity: r[7], Profit: r[8], "Shipping Days": r[9]
}));

const chartColors = ["#20d9ff", "#7c4dff", "#39e5a1", "#2d7cff", "#b84dff", "#18b5e8"];

function App() {
  const [dataset, setDataset] = useState<Dataset | null>(null);
  const [page, setPage] = useState<Page>("landing");
  const [filters, setFilters] = useState<Filters>({ date: "", region: "", segment: "", category: "" });
  const [dragging, setDragging] = useState(false);
  const [busy, setBusy] = useState(false);
  const [sidebar, setSidebar] = useState(true);
  const [companyName, setCompanyName] = useState(() => {
    try { return localStorage.getItem("sbi-company") || "Your Company"; } catch { return "Your Company"; }
  });
  const [reportOpen, setReportOpen] = useState(false);
  const [settingsOpen, setSettingsOpen] = useState(false);
  const [theme, setTheme] = useState<"light" | "dark">(() => {
    try {
      return (localStorage.getItem("sbi-theme") as "light" | "dark") || "light";
    } catch {
      return "light";
    }
  });

  useEffect(() => {
    try { localStorage.setItem("sbi-company", companyName); } catch {}
  }, [companyName]);

  useEffect(() => {
    document.documentElement.dataset.theme = theme;
    try { localStorage.setItem("sbi-theme", theme); } catch {}
  }, [theme]);
  const fileRef = useRef<HTMLInputElement>(null);

  const load = async (files: FileList | File[]) => {
    const first = Array.from(files)[0];
    if (!first) return;
    setBusy(true);
    try {
      const parsed = await parseFile(first);
      if (!parsed.rows.length) throw new Error("No tabular rows were found.");
      setDataset(parsed);
      setFilters({ date: "", region: "", segment: "", category: "" });
      setPage("overview");
    } catch (e) {
      alert(e instanceof Error ? e.message : "Could not read this file.");
    } finally {
      setBusy(false);
    }
  };

  const loadDemo = () => {
    setDataset(makeDataset("demo-sales-data.csv", "demo", demoRows));
    setFilters({ date: "", region: "", segment: "", category: "" });
    setPage("overview");
  };

  const rows = useMemo(() => dataset ? applyFilters(dataset, filters) : [], [dataset, filters]);

  const years = useMemo(() => {
    if (!dataset) return [];
    const col = findColumn(dataset, ["Order Date", "Date", "Ship Date", "OrderDate"]);
    if (!col) return [];
    return [...new Set(dataset.rows.map(r => new Date(String(r[col])).getFullYear()).filter(y => Number.isFinite(y) && y > 1900))].sort();
  }, [dataset]);

  const filterOptions = useMemo(() => {
    if (!dataset) return { regions: [], segments: [], categories: [] };
    const m = inferMetrics(dataset);
    return {
      regions: uniqueValues(dataset.rows, m.region),
      segments: uniqueValues(dataset.rows, m.segment),
      categories: uniqueValues(dataset.rows, m.category)
    };
  }, [dataset]);

  const reset = () => setFilters({ date: "", region: "", segment: "", category: "" });

  return (
    <div className="app-shell">
      <header className="topbar">
        <button
            className="theme-toggle"
            type="button"
            title={theme === "light" ? "Switch to dark theme" : "Switch to light theme"}
            onClick={() => setTheme((t) => (t === "light" ? "dark" : "light"))}
          >
            {theme === "light" ? <Moon size={17} /> : <Sun size={17} />}
            <span>{theme === "light" ? "Dark" : "Light"}</span>
          </button>
        <div className="brand">
          <div className="brand-mark"><BarChart3 size={22} /></div>
          <span>Sales Business Insights</span>
        </div>
        <nav className="top-nav">
          {(["landing", "overview", "product", "regional", "insights"] as Page[]).map(p => (
            <button key={p} className={page === p ? "nav-item active" : "nav-item"} onClick={() => dataset || p === "landing" ? setPage(p) : undefined}>
              {p === "landing" ? <Home size={17}/> : p === "overview" ? <LayoutDashboard size={17}/> : p === "product" ? <ShoppingCart size={17}/> : p === "regional" ? <Globe2 size={17}/> : <Sparkles size={17}/>}
              {pageTitles[p]}
            </button>
          ))}
        </nav>
        <div className="top-actions">
          {dataset && <span className="file-pill"><Database size={14}/> {dataset.name}</span>}
          {dataset && <button className="report-btn" title="Download report" onClick={() => setReportOpen(v => !v)}><FileText size={16}/> Report</button>}
          <button className="icon-btn" title="Company & report settings" onClick={() => setSettingsOpen(true)}><Settings2 size={18}/></button>
          <button className="icon-btn" title="Upload a new dataset" onClick={() => fileRef.current?.click()}><UploadCloud size={18}/></button>
          <button className="icon-btn" title="Refresh" onClick={() => dataset && setDataset({...dataset})}><RefreshCw size={17}/></button>
        </div>
        {dataset && reportOpen && (
          <ReportMenu
            dataset={dataset}
            rows={rows}
            page={page}
            companyName={companyName}
            onClose={() => setReportOpen(false)}
          />
        )}
        {settingsOpen && (
          <SettingsModal
            companyName={companyName}
            setCompanyName={setCompanyName}
            onClose={() => setSettingsOpen(false)}
          />
        )}
        <input ref={fileRef} hidden type="file" accept=".csv,.xlsx,.xls,.txt,.sql" onChange={e => e.target.files && load(e.target.files)} />
      </header>

      <aside className={sidebar ? "sidebar" : "sidebar collapsed"}>
        <button className="collapse-btn" onClick={() => setSidebar(v => !v)}>{sidebar ? <ChevronLeft size={18}/> : <ChevronRight size={18}/>}</button>
        <div className="side-links">
          <SideButton icon={<Home/>} active={page === "landing"} label="Landing" open={sidebar} onClick={() => setPage("landing")}/>
          <SideButton icon={<BarChart3/>} active={page === "overview"} label="Overview" open={sidebar} onClick={() => dataset && setPage("overview")}/>
          <SideButton icon={<Box/>} active={page === "product"} label="Products" open={sidebar} onClick={() => dataset && setPage("product")}/>
          <SideButton icon={<Truck/>} active={page === "regional"} label="Operations" open={sidebar} onClick={() => dataset && setPage("regional")}/>
          <SideButton icon={<Users/>} label="Customers" open={sidebar}/>
          <SideButton icon={<Sparkles/>} active={page === "insights"} label="Key Insights" open={sidebar} onClick={() => dataset && setPage("insights")}/>
          <SideButton icon={<Settings2/>} label="Settings" open={sidebar}/>
        </div>
        <div className="side-bottom"><Zap size={16}/>{sidebar && <span>Auto insights</span>}</div>
      </aside>

      <main className={sidebar ? "content" : "content wide"}>
        {page === "landing" || !dataset
          ? <Landing busy={busy} onPick={() => fileRef.current?.click()} onDrop={load} onDemo={loadDemo} />
          : <Dashboard dataset={dataset} rows={rows} page={page} filters={filters} setFilters={setFilters} reset={reset} years={years} options={filterOptions} />}
      </main>
    </div>
  );
}

function ReportMenu({dataset,rows,page,companyName,onClose}:{dataset:Dataset;rows:Record<string,unknown>[];page:Page;companyName:string;onClose:()=>void}) {
  const exportData = buildReportData(dataset, rows, companyName, page);
  return <div className="report-menu">
    <div className="report-menu-head"><div><b>Download Report</b><span>Executive-ready export from current filters</span></div><button className="mini-close" onClick={onClose}><X size={15}/></button></div>
    <button onClick={() => {downloadExcel(exportData);onClose()}}><FileSpreadsheet size={17}/><span><b>Excel Report</b><small>Summary + insights + filtered data</small></span></button>
    <button onClick={() => {downloadPdf(exportData);onClose()}}><FileText size={17}/><span><b>PDF Report</b><small>Executive summary for sharing</small></span></button>
    <button onClick={() => {downloadPpt(exportData);onClose()}}><LayoutDashboard size={17}/><span><b>PowerPoint Report</b><small>Presentation-ready slides</small></span></button>
  </div>;
}

function SettingsModal({companyName,setCompanyName,onClose}:{companyName:string;setCompanyName:(v:string)=>void;onClose:()=>void}) {
  const [value,setValue]=useState(companyName);
  return <div className="modal-backdrop" onMouseDown={e=>{if(e.target===e.currentTarget)onClose()}}><div className="settings-modal">
    <div className="modal-head"><div><b>Company & Report Settings</b><span>This name appears on exported PDF, Excel and PowerPoint reports.</span></div><button className="mini-close" onClick={onClose}><X size={16}/></button></div>
    <label className="settings-field"><span>Company Name</span><input autoFocus value={value} onChange={e=>setValue(e.target.value)} placeholder="Enter company name" /></label>
    <div className="modal-actions"><button className="secondary-btn" onClick={onClose}>Cancel</button><button className="primary-btn" onClick={()=>{setCompanyName(value.trim()||"Your Company");onClose()}}>Save Company Name</button></div>
  </div></div>;
}

type ReportData = {
  company:string; title:string; dataset:string; generated:string; rows:Record<string,unknown>[];
  sales:number; profit:number; margin:number; orders:number; customers:number; lossRows:number; lossValue:number;
  growth:number|null; topRegion:string; topCategory:string; topProduct:string;
  region:{name:string;value:number}[]; category:{name:string;value:number}[]; monthly:{name:string;value:number}[];
};

function buildReportData(dataset:Dataset, rows:Record<string,unknown>[], company:string, page:Page):ReportData {
  const m=inferMetrics(dataset), sales=sumBy(rows,m.sales), profit=sumBy(rows,m.profit), margin=sales?profit/sales*100:0;
  const orders=m.orderId?new Set(rows.map(r=>String(r[m.orderId]))).size:rows.length;
  const customers=m.customer?new Set(rows.map(r=>String(r[m.customer]))).size:0;
  const lossRows=m.profit?rows.filter(r=>toReportNumber(r[m.profit])<0).length:0;
  const lossValue=m.profit?Math.abs(rows.reduce((a,r)=>{const n=toReportNumber(r[m.profit]);return n<0?a+n:a},0)):0;
  const monthly = trend(rows, m.date, m.sales); const last = monthly[monthly.length - 1]; const prev = monthly[monthly.length - 2];
  const growth=prev&&prev.value!==0?((last!.value-prev.value)/prev.value)*100:null;
  return {company:company||"Your Company",title:pageTitles[page],dataset:dataset.name,generated:new Date().toLocaleString(),rows,sales,profit,margin,orders,customers,lossRows,lossValue,growth,
    topRegion:groupSum(rows,m.region,m.sales)[0]?.name||"Not detected",topCategory:groupSum(rows,m.category,m.sales)[0]?.name||"Not detected",topProduct:groupSum(rows,m.product,m.sales)[0]?.name||"Not detected",
    region:groupSum(rows,m.region,m.sales,8),category:groupSum(rows,m.category,m.sales,8),monthly:monthly.slice(-12)};
}
function toReportNumber(v:unknown){return Number(String(v??"").replace(/[$,%\s,]/g,""))||0}
function safeFile(s:string){return s.replace(/[^a-z0-9-_]+/gi,"-").replace(/-+/g,"-").replace(/^-|-$/g,"")||"report"}
function downloadExcel(r:ReportData){
  const wb=XLSX.utils.book_new();
  const summary=[
    [r.company,"Sales Business Insights Report"],["Generated",r.generated],["Dataset",r.dataset],["Report View",r.title],[],
    ["KPI","Value"],["Total Sales",r.sales],["Net Profit",r.profit],["Profit Margin %",r.margin],["Orders",r.orders],["Customers",r.customers],["Loss-making Records",r.lossRows],["Absolute Loss",r.lossValue],["Revenue Growth %",r.growth??"Not available"],[],
    ["Top Region",r.topRegion],["Top Category",r.topCategory],["Top Product",r.topProduct]
  ];
  XLSX.utils.book_append_sheet(wb,XLSX.utils.aoa_to_sheet(summary),"Executive Summary");
  XLSX.utils.book_append_sheet(wb,XLSX.utils.json_to_sheet(r.rows),"Filtered Data");
  XLSX.utils.book_append_sheet(wb,XLSX.utils.json_to_sheet(r.region),"Region Performance");
  XLSX.utils.book_append_sheet(wb,XLSX.utils.json_to_sheet(r.category),"Category Performance");
  XLSX.utils.book_append_sheet(wb,XLSX.utils.json_to_sheet(r.monthly),"Monthly Trend");
  XLSX.writeFile(wb,`${safeFile(r.company)}-Sales-Business-Report.xlsx`);
}
function downloadPdf(r:ReportData){
  const pdf=new jsPDF({unit:"pt",format:"a4"}); const left=42; let y=48;
  pdf.setFontSize(20); pdf.text(r.company,left,y); y+=25; pdf.setFontSize(14); pdf.text("Sales Business Insights — Executive Report",left,y); y+=18; pdf.setFontSize(9); pdf.text(`Generated: ${r.generated}  |  Dataset: ${r.dataset}`,left,y); y+=30;
  const cards=[[`Total Sales`,currency(r.sales)],[`Net Profit`,currency(r.profit)],[`Margin`,`${r.margin.toFixed(2)}%`],[`Orders`,r.orders.toLocaleString()],[`Customers`,r.customers.toLocaleString()],[`Loss Records`,r.lossRows.toLocaleString()]];
  pdf.setFontSize(10); cards.forEach((c,i)=>{const x=left+(i%2)*260; if(i%2===0&&i>0)y+=48; pdf.rect(x,y-15,245,38); pdf.text(c[0],x+10,y); pdf.setFontSize(14); pdf.text(c[1],x+10,y+17); pdf.setFontSize(10)}); y+=65;
  pdf.setFontSize(13); pdf.text("Business Summary",left,y); y+=20; pdf.setFontSize(10);
  const lines=[`Revenue growth: ${r.growth===null?"Not available":`${r.growth>=0?"+":""}${r.growth.toFixed(1)}%`}`,`Top region: ${r.topRegion}`,`Top category: ${r.topCategory}`,`Top product: ${r.topProduct}`,`Absolute loss: ${currency(r.lossValue)}`]; lines.forEach(t=>{pdf.text(`• ${t}`,left,y);y+=16}); y+=12;
  pdf.setFontSize(13);pdf.text("Top Region Performance",left,y);y+=20;pdf.setFontSize(9);r.region.slice(0,8).forEach((d,i)=>{pdf.text(`${i+1}. ${d.name}`,left,y);pdf.text(currency(d.value),430,y);y+=15});
  if(y>735){pdf.addPage();y=48} y+=15;pdf.setFontSize(13);pdf.text("Top Categories",left,y);y+=20;pdf.setFontSize(9);r.category.slice(0,8).forEach((d,i)=>{pdf.text(`${i+1}. ${d.name}`,left,y);pdf.text(currency(d.value),430,y);y+=15});
  pdf.setFontSize(8);pdf.text("Generated from uploaded data. Figures are decision-support outputs and should be validated before financial use.",left,805);
  pdf.save(`${safeFile(r.company)}-Sales-Business-Report.pdf`);
}
function downloadPpt(r:ReportData){
  const ppt=new PptxGenJS(); ppt.layout="LAYOUT_WIDE"; ppt.author=r.company; ppt.subject="Sales Business Insights"; ppt.title=`${r.company} Sales Business Report`;
  const addTitle=(slide:any,title:string,sub:string)=>{slide.addText(r.company,{x:.5,y:.25,w:12,h:.35,fontSize:20,bold:true,color:"1B2A41"});slide.addText(title,{x:.5,y:.8,w:12,h:.5,fontSize:26,bold:true,color:"172033"});slide.addText(sub,{x:.5,y:1.35,w:12,h:.3,fontSize:10,color:"66758A"})};
  let slide=ppt.addSlide();addTitle(slide,"Executive Business Report",`Dataset: ${r.dataset} · ${r.generated}`); const vals=[[`Total Sales`,currency(r.sales)],[`Net Profit`,currency(r.profit)],[`Margin`,`${r.margin.toFixed(2)}%`],[`Orders`,r.orders.toLocaleString()]]; vals.forEach((v,i)=>slide.addText(`${v[0]}\n${v[1]}`,{x:.7+i*3.05,y:2.2,w:2.7,h:1.15,fontSize:16,bold:true,color:"172033",fill:{color:"F4F7FB"},line:{color:"D9E2EF"},margin:.18,breakLine:false,fit:"shrink"}));
  slide.addText(`Revenue Growth: ${r.growth===null?"Not available":`${r.growth>=0?"+":""}${r.growth.toFixed(1)}%`}\nTop Region: ${r.topRegion}\nTop Category: ${r.topCategory}\nTop Product: ${r.topProduct}\nLoss-making Records: ${r.lossRows.toLocaleString()}`,{x:.8,y:4.1,w:11.3,h:2.1,fontSize:18,color:"172033",breakLine:false,fit:"shrink"});
  slide=ppt.addSlide();addTitle(slide,"Regional Performance","Sales contribution by leading geography"); slide.addChart(ppt.ChartType.bar,[{name:"Sales",labels:r.region.map(d=>d.name),values:r.region.map(d=>d.value)}],{x:.7,y:1.9,w:11.8,h:4.6,catAxisLabelFontSize:9,valAxisLabelFontSize:9,showLegend:false,showTitle:false});
  slide=ppt.addSlide();addTitle(slide,"Category Performance","Revenue contribution by category"); slide.addChart(ppt.ChartType.bar,[{name:"Sales",labels:r.category.map(d=>d.name),values:r.category.map(d=>d.value)}],{x:.7,y:1.9,w:11.8,h:4.6,catAxisLabelFontSize:9,valAxisLabelFontSize:9,showLegend:false});
  slide=ppt.addSlide();addTitle(slide,"Management Takeaways","Automatic decision-support summary from the uploaded data");slide.addText(`1. ${r.topRegion} is the leading region by sales.\n2. ${r.topCategory} is the leading category by revenue.\n3. ${r.topProduct} is the leading product by revenue.\n4. ${r.lossRows.toLocaleString()} records are loss-making, representing ${currency(r.lossValue)} absolute loss.\n5. Validate pricing, discount, freight and product/customer mix before scaling low-margin revenue.`,{x:.8,y:2,w:11,h:3.8,fontSize:18,color:"172033",breakLine:false,fit:"shrink"});
  ppt.writeFile({fileName:`${safeFile(r.company)}-Sales-Business-Report.pptx`});
}

function SideButton({icon, label, active, open, onClick}: {icon: React.ReactNode; label: string; active?: boolean; open: boolean; onClick?: ()=>void}) {
  return <button className={active ? "side-link active" : "side-link"} onClick={onClick}>{icon}{open && <span>{label}</span>}</button>;
}

function Landing({busy, onPick, onDrop, onDemo}: {busy:boolean; onPick:()=>void; onDrop:(f:FileList)=>void; onDemo:()=>void}) {
  return (
    <section className="landing-page">
      <div className="hero-copy">
        <div className="eyebrow">EXECUTIVE PROJECT OVERVIEW <i/> DATA → INSIGHTS</div>
        <h1>Sales <span>Business</span> Insights</h1>
        <p>A unified analytics workspace that turns raw Excel, CSV, TXT and SQL files into executive-ready interactive reports.</p>
      </div>
      <div className="landing-grid">
        <div className="upload-card">
          <div className="upload-icon"><UploadCloud size={34}/></div>
          <h2>{busy ? "Building your report…" : "Drop your data here"}</h2>
          <p>Upload a file and the dashboard will automatically profile columns, detect metrics, create filters, and generate visual stories.</p>
          <div
            className="drop-zone"
            onDragOver={e => { e.preventDefault(); }}
            onDragEnter={() => {}}
            onDrop={e => { e.preventDefault(); if (e.dataTransfer.files.length) onDrop(e.dataTransfer.files); }}
          >
            <FileSpreadsheet/><FileText/>
            <span>Excel · CSV · TXT · SQL</span>
          </div>
          <button className="primary-btn" onClick={onPick} disabled={busy}><UploadCloud size={18}/> {busy ? "Analyzing…" : "Choose file"}</button>
          <button className="secondary-btn" onClick={onDemo}><Sparkles size={17}/> Preview with demo data</button>
        </div>
        <div className="landing-art">
          <div className="art-ring r1"/><div className="art-ring r2"/>
          <div className="floating-kpi"><CircleDollarSign/><strong>$1.12M</strong><small>TOTAL SALES</small></div>
          <div className="floating-kpi green"><Target/><strong>$392.7K</strong><small>NET PROFIT</small></div>
          <div className="bars-art"><span/><span/><span/><span/><span/></div>
          <div className="orbit-line"/>
        </div>
      </div>
      <div className="feature-row">
        <Feature icon={<Table2/>} title="Auto profiling" text="Types, nulls, unique values & metrics"/>
        <Feature icon={<Layers3/>} title="Auto dashboard" text="KPI cards, charts, filters & insights"/>
        <Feature icon={<Globe2/>} title="Interactive views" text="Executive, product & regional pages"/>
        <Feature icon={<Zap/>} title="100% browser-side" text="Your files stay in this app"/>
      </div>
    </section>
  );
}

function Feature({icon,title,text}:{icon:React.ReactNode;title:string;text:string}) {
  return <div className="feature"><div className="feature-icon">{icon}</div><div><b>{title}</b><span>{text}</span></div></div>;
}

function Dashboard({
  dataset, rows, page, filters, setFilters, reset, years, options
}: {
  dataset: Dataset; rows: Record<string, unknown>[]; page: Page; filters: Filters; setFilters: React.Dispatch<React.SetStateAction<Filters>>;
  reset: ()=>void; years: number[]; options: {regions:string[];segments:string[];categories:string[]};
}) {
  const m = inferMetrics(dataset);
  const sales = sumBy(rows, m.sales);
  const profit = sumBy(rows, m.profit);
  const qty = sumBy(rows, m.quantity);
  const margin = sales ? profit / sales * 100 : 0;
  const shipping = avgBy(rows, m.shipping);
  const orders = m.orderId ? new Set(rows.map(r => String(r[m.orderId]))).size : rows.length;
  const customers = m.customer ? new Set(rows.map(r => String(r[m.customer]))).size : 0;
  const states = m.region ? new Set(rows.map(r => String(r[m.region]))).size : 0;

  const subtitle =
    page === "overview" ? "Sales, Profitability & Trend Overview" :
    page === "product" ? "Category, Product & Customer Performance" :
    page === "regional" ? "Geography, Order Flow & Fulfillment Performance" :
    "Growth, Profitability, Loss Analysis & Recommended Actions";

  return (
    <>
      <div className="page-heading">
        <div>
          <div className="eyebrow">DATASET · {dataset.sourceType.toUpperCase()} <i/> {rows.length.toLocaleString()} ROWS</div>
          <h1>{pageTitles[page]}</h1>
          <p>{subtitle}</p>
        </div>
        <div className="heading-actions"><button className="ghost-btn" onClick={reset}>Reset filters</button><span className="live-dot"/> Live analysis</div>
      </div>

      {page === "overview" && <Overview dataset={dataset} rows={rows} m={m} sales={sales} profit={profit} margin={margin} shipping={shipping} orders={orders} />}
      {page === "product" && <Product dataset={dataset} rows={rows} m={m} sales={sales} customers={customers} />}
      {page === "regional" && <Regional dataset={dataset} rows={rows} m={m} orders={orders} qty={qty} states={states} shipping={shipping} />}
      {page === "insights" && <BusinessInsights dataset={dataset} rows={rows} m={m} sales={sales} profit={profit} margin={margin} orders={orders} customers={customers} />}

      <div className="filter-bar">
        <Filter label="Year" value={filters.date} options={years.map(String)} onChange={v => setFilters(f => ({...f,date:v}))}/>
        <Filter label="Region / Geography" value={filters.region} options={options.regions} onChange={v => setFilters(f => ({...f,region:v}))}/>
        <Filter label="Client Segment" value={filters.segment} options={options.segments} onChange={v => setFilters(f => ({...f,segment:v}))}/>
        <Filter label="Product Category" value={filters.category} options={options.categories} onChange={v => setFilters(f => ({...f,category:v}))}/>
      </div>

      <IntelligenceStrip dataset={dataset} rows={rows} />
    </>
  );
}


function BusinessInsights({
  dataset, rows, m, sales, profit, margin, orders, customers
}: {
  dataset: Dataset; rows: Record<string, unknown>[]; m: ReturnType<typeof inferMetrics>;
  sales: number; profit: number; margin: number; orders: number; customers: number;
}) {
  const monthly = trend(rows, m.date, m.sales);
  const last = monthly[monthly.length - 1];
const previous = monthly[monthly.length - 2];
  const growth = previous && previous.value !== 0 ? ((last!.value - previous.value) / previous.value) * 100 : null;

  const profitRows = m.profit ? rows.filter(r => Number(String(r[m.profit] ?? "").replace(/[$,%\s,]/g, "")) > 0).length : 0;
  const lossRows = m.profit ? rows.filter(r => Number(String(r[m.profit] ?? "").replace(/[$,%\s,]/g, "")) < 0).length : 0;
  const lossValue = m.profit ? Math.abs(rows.reduce((sum, r) => {
    const n = Number(String(r[m.profit] ?? "").replace(/[$,%\s,]/g, ""));
    return Number.isFinite(n) && n < 0 ? sum + n : sum;
  }, 0)) : 0;

  const regions = groupSum(rows, m.region, m.sales, 6);
  const categories = groupSum(rows, m.category, m.sales, 6);
  const products = groupSum(rows, m.product, m.sales, 6);
  const customerSales = groupSum(rows, m.customer, m.sales, 6);

  const topRegion = regions[0];
  const topCategory = categories[0];
  const topProduct = products[0];

  const health = Math.max(0, Math.min(100,
    50 + Math.min(25, Math.max(-25, (growth ?? 0) * 0.7)) +
    Math.min(20, Math.max(-20, margin - 10)) -
    Math.min(15, (lossRows / Math.max(rows.length, 1)) * 100)
  ));

  const healthLabel = health >= 75 ? "Strong" : health >= 55 ? "Stable" : health >= 40 ? "Watch" : "At Risk";
  const growthPositive = growth === null || growth >= 0;

  const actions = [
    topRegion ? {
      icon: <ArrowUpRight/>,
      tone: "positive",
      title: `${topRegion.name} is your strongest geography`,
      body: `${topRegion.name} contributes ${sales ? ((topRegion.value / sales) * 100).toFixed(1) : 0}% of sales. Protect availability and replicate its winning approach in weaker markets.`
    } : null,
    topCategory ? {
      icon: <Target/>,
      tone: "neutral",
      title: `${topCategory.name} leads category revenue`,
      body: `Use ${topCategory.name} as the growth benchmark, but prioritize its profitable products and customers rather than revenue alone.`
    } : null,
    lossRows > 0 ? {
      icon: <AlertTriangle/>,
      tone: "risk",
      title: `${lossRows.toLocaleString()} records are loss-making`,
      body: `Negative profit totals at least ${currency(lossValue)} in absolute loss. Investigate discounts, cost, freight and low-margin products behind these transactions.`
    } : {
      icon: <CheckCircle2/>,
      tone: "positive",
      title: "No negative-profit records detected",
      body: "Every analyzed transaction has non-negative profit. Continue monitoring margin by product and customer."
    }
  ].filter(Boolean) as {icon: React.ReactNode; tone:string; title:string; body:string}[];

  return (
    <section className="insights-page">
      <div className="insights-nav-card">
        <div className="nav-copy">
          <b>Business Decision Center</b>
          <span>Growth · Profitability · Risk · Opportunities</span>
        </div>
        <div className="insight-tabs">
          <button className="insight-tab active">Overview</button>
          <button className="insight-tab" onClick={() => window.scrollTo({top: 520, behavior: "smooth"})}>Growth</button>
          <button className="insight-tab" onClick={() => window.scrollTo({top: 760, behavior: "smooth"})}>Risks</button>
        </div>
      </div>
      <div className="insights-hero">
        <div>
          <div className="eyebrow">BUSINESS INTELLIGENCE <i/> DECISION SUPPORT</div>
          <h2>What is happening in your business?</h2>
          <p>Automatic analysis of growth, profitability, customers, products and operational risk from the uploaded dataset.</p>
        </div>
        <div className={`health-score ${healthLabel.toLowerCase().replace(" ","-")}`}>
          <span>Business Health</span>
          <strong>{Math.round(health)}<small>/100</small></strong>
          <b>{healthLabel}</b>
        </div>
      </div>

      <div className="insight-kpis">
        <div className="insight-kpi"><span>Revenue Growth</span><strong className={growthPositive ? "positive-text" : "negative-text"}>{growth === null ? "—" : `${growth >= 0 ? "+" : ""}${growth.toFixed(1)}%`}</strong><small>Latest month vs previous month</small></div>
        <div className="insight-kpi"><span>Net Profit</span><strong>{currency(profit)}</strong><small>{margin.toFixed(2)}% overall margin</small></div>
        <div className="insight-kpi"><span>Loss-making Records</span><strong className={lossRows ? "negative-text" : "positive-text"}>{lossRows.toLocaleString()}</strong><small>{rows.length ? ((lossRows / rows.length) * 100).toFixed(1) : 0}% of analyzed rows</small></div>
        <div className="insight-kpi"><span>Profit / Order</span><strong>{orders ? currency(profit / orders) : "—"}</strong><small>{customers.toLocaleString()} customers analyzed</small></div>
      </div>

      <div className="insight-layout">
        <Panel title="Growth & Profit Story" className="insight-main-panel">
          <div className="story-grid">
            <div className="story-card">
              <div className="story-icon cyan">{growthPositive ? <TrendingUp/> : <TrendingDown/>}</div>
              <div>
                <span>Revenue direction</span>
                <strong>{growth === null ? "Not enough time data" : growthPositive ? "Growing" : "Declining"}</strong>
                <p>{growth === null ? "Add a valid date field to calculate period-over-period growth." : `Revenue moved from ${currency(previous!.value)} to ${currency(last!.value)} in the latest two periods.`}</p>
              </div>
            </div>
            <div className="story-card">
              <div className="story-icon green"><Target/></div>
              <div>
                <span>Profitability</span>
                <strong>{margin >= 20 ? "Healthy margin" : margin >= 10 ? "Moderate margin" : "Thin margin"}</strong>
                <p>Current net margin is {margin.toFixed(2)}%. Focus on high-margin categories, customers and products before chasing low-quality revenue.</p>
              </div>
            </div>
          </div>
          <div className="mini-metrics">
            <div><span>Top Region</span><b>{topRegion?.name ?? "—"}</b><em>{topRegion ? currency(topRegion.value) : "—"}</em></div>
            <div><span>Top Category</span><b>{topCategory?.name ?? "—"}</b><em>{topCategory ? currency(topCategory.value) : "—"}</em></div>
            <div><span>Top Product</span><b>{topProduct?.name ?? "—"}</b><em>{topProduct ? currency(topProduct.value) : "—"}</em></div>
            <div><span>Orders / Customers</span><b>{orders.toLocaleString()}</b><em>{customers.toLocaleString()} customers</em></div>
          </div>
        </Panel>

        <Panel title="Recommended Actions" className="actions-panel">
          {actions.map((a, i) => <div className={`action-card ${a.tone}`} key={i}><div className="action-icon">{a.icon}</div><div><b>{a.title}</b><p>{a.body}</p></div></div>)}
        </Panel>
      </div>

      <div className="insight-layout lower">
        <Panel title="Top Customers by Revenue"><BarHorizontal data={customerSales}/></Panel>
        <Panel title="Revenue by Category"><BarVertical data={categories}/></Panel>
        <Panel title="Risk Signals">
          <div className="risk-list">
            <div className={lossRows ? "risk danger" : "risk good"}><span>{lossRows ? "Loss-making transactions" : "Profitability"}</span><strong>{lossRows ? lossRows.toLocaleString() : "Positive"}</strong><small>{lossRows ? `${currency(lossValue)} absolute loss` : "No negative-profit rows detected"}</small></div>
            <div className={margin < 10 ? "risk danger" : "risk good"}><span>Net margin</span><strong>{margin.toFixed(2)}%</strong><small>{margin < 10 ? "Below 10% — review pricing/costs" : "Above 10% — continue margin monitoring"}</small></div>
            <div className="risk good"><span>Data completeness</span><strong>{dataset.quality?.completeness ?? 100}%</strong><small>{dataset.quality?.duplicateRows ?? 0} duplicate rows detected</small></div>
          </div>
        </Panel>
      </div>

      <div className="insight-disclaimer">
        <Sparkles size={14}/> Insights are calculated from the uploaded data and are decision-support signals, not financial advice.
      </div>
    </section>
  );
}

function Filter({label,value,options,onChange}:{label:string;value:string;options:string[];onChange:(v:string)=>void}) {
  return <label className="filter"><span>{label}</span><select value={value} onChange={e=>onChange(e.target.value)}><option value="">All</option>{options.map(v=><option key={v}>{v}</option>)}</select></label>;
}

function IntelligenceStrip({dataset, rows}: {dataset: Dataset; rows: Record<string, unknown>[]}) {
  const m = inferMetrics(dataset);
  const q = dataset.quality;
  const detected = [
    ["Sales", m.sales],
    ["Profit", m.profit],
    ["Date", m.date],
    ["Geography", m.region],
    ["Category", m.category],
    ["Product", m.product],
    ["Customer", m.customer],
    ["Shipping", m.shipping ?? (m.shipDate && m.date ? "Calculated from dates" : undefined)]
  ];
  const found = detected.filter(([, value]) => Boolean(value)).length;
  const missing = detected.length - found;
  return (
    <div className="intelligence-strip">
      <div className="intel-summary">
        <div className="intel-badge"><Sparkles size={15}/></div>
        <div>
          <b>Data Intelligence</b>
          <span>Automatic schema, metric & quality analysis</span>
        </div>
      </div>
      <div className="intel-mappings">
        {detected.slice(0, 6).map(([label, value]) => (
          <div className="intel-chip" key={label}>
            <span>{label}</span>
            <strong title={String(value ?? "Not detected")}>{value ? String(value) : "Not detected"}</strong>
          </div>
        ))}
      </div>
      <div className="intel-quality">
        <div><span>Mapped</span><strong>{found}/{detected.length}</strong></div>
        <div><span>Completeness</span><strong>{q ? `${q.completeness}%` : "—"}</strong></div>
        <div><span>Rows</span><strong>{rows.length.toLocaleString()}</strong></div>
        {missing > 0 && <div className="warning"><span>Needs mapping</span><strong>{missing}</strong></div>}
      </div>
    </div>
  );
}

function Kpi({label,value,accent="cyan",icon}:{label:string;value:string;accent?:string;icon:React.ReactNode}) {
  return <div className={`kpi ${accent}`}><div className="kpi-icon">{icon}</div><div><span>{label}</span><strong>{value}</strong></div></div>;
}

function Panel({title,children,className=""}:{title:string;children:React.ReactNode;className?:string}) {
  return <div className={`panel ${className}`}><div className="panel-title">{title}</div>{children}</div>;
}

function Overview({dataset,rows,m,sales,profit,margin,shipping,orders}:{dataset:Dataset;rows:Record<string,unknown>[];m:ReturnType<typeof inferMetrics>;sales:number;profit:number;margin:number;shipping:number;orders:number}) {
  const region = groupSum(rows,m.region,m.sales);
  const category = groupSum(rows,m.category,m.sales);
  const segment = groupSum(rows,m.segment,m.sales);
  const tr = trend(rows,m.date,m.sales);
  const scatter = m.sales && m.profit && m.region ? groupScatter(rows,m.region,m.sales,m.profit) : [];
  return <>
    <div className="kpi-grid">
      <Kpi label="Total Sales" value={currency(sales)} icon={<CircleDollarSign/>}/>
      <Kpi label="Net Profit" value={currency(profit)} accent="green" icon={<Target/>}/>
      <Kpi label="Profit Margin %" value={`${margin.toFixed(2)}%`} accent="purple" icon={<Target/>}/>
      <Kpi label="Average Shipping Days" value={shipping > 0 ? `${shipping.toFixed(1)} Days` : "Not detected"} accent="blue" icon={<Truck/>}/>
    </div>
    <div className="chart-grid top">
      <Panel title="Total Sales & Net Profit Trend (Monthly)" className="wide-panel"><LineTrend data={tr} profit={m.profit ? trend(rows,m.date,m.profit) : []}/></Panel>
      <Panel title="Sales by Region"><BarHorizontal data={region}/></Panel>
      <Panel title="Sales by Product Category"><BarVertical data={category}/></Panel>
    </div>
    <div className="chart-grid bottom">
      <Panel title="Sales by Client Segment"><BarHorizontal data={segment}/></Panel>
      <Panel title="Performance: Sales vs Net Profit" className="wide-panel"><ScatterPlot data={scatter}/></Panel>
      <InsightPanel dataset={dataset} rows={rows} sales={sales} profit={profit} orders={orders}/>
    </div>
  </>;
}

function Product({dataset,rows,m,sales,customers}:{dataset:Dataset;rows:Record<string,unknown>[];m:ReturnType<typeof inferMetrics>;sales:number;customers:number}) {
  const category = groupSum(rows,m.category,m.sales);
  const product = groupSum(rows,m.product,m.sales);
  const customer = groupSum(rows,m.customer,m.sales);
  const scatter = m.sales && m.profit && m.product ? groupScatter(rows,m.product,m.sales,m.profit) : [];
  return <>
    <div className="kpi-grid">
      <Kpi label="Product Entities" value={number(m.product ? new Set(rows.map(r=>String(r[m.product]))).size : dataset.rows.length)} icon={<Box/>}/>
      <Kpi label="Total Customers" value={number(customers)} accent="green" icon={<Users/>}/>
      <Kpi label="Top Category Sales" value={category.length ? currency(category[0].value) : "Not detected"} accent="purple" icon={<Target/>}/>
      <Kpi label="Top Product Sales" value={product.length ? currency(product[0].value) : "Not detected"} accent="blue" icon={<ShoppingCart/>}/>
    </div>
    <div className="chart-grid top">
      <Panel title="Product Category Contribution (Sales)"><BarVertical data={category}/></Panel>
      <Panel title="Top Products by Sales"><BarHorizontal data={product}/></Panel>
      <Panel title="Sales vs Net Profit by Product"><ScatterPlot data={scatter}/></Panel>
    </div>
    <div className="chart-grid bottom">
      <Panel title="Sales by Client Segment"><BarHorizontal data={groupSum(rows,m.segment,m.sales)}/></Panel>
      <Panel title="Top Customers by Sales" className="wide-panel"><BarHorizontal data={customer}/></Panel>
      <InsightPanel dataset={dataset} rows={rows} sales={sales} profit={sumBy(rows,m.profit)} orders={rows.length}/>
    </div>
  </>;
}

function Regional({dataset,rows,m,orders,qty,states,shipping}:{dataset:Dataset;rows:Record<string,unknown>[];m:ReturnType<typeof inferMetrics>;orders:number;qty:number;states:number;shipping:number}) {
  const region = groupSum(rows,m.region,m.sales);
  const quarter = trend(rows,m.date,m.quantity);
  const segment = groupSum(rows,m.segment);
  const scatter = m.sales && m.shipping && m.region ? groupScatter(rows,m.region,m.sales,m.shipping) : [];
  return <>
    <div className="kpi-grid">
      <Kpi label="Total Orders" value={number(orders)} icon={<ShoppingCart/>}/>
      <Kpi label="Total Units" value={number(qty)} accent="green" icon={<Package/>}/>
      <Kpi label="States / Regions Covered" value={number(states)} accent="purple" icon={<Globe2/>}/>
      <Kpi label="Average Order Value" value={orders && m.sales ? currency(sumBy(rows,m.sales)/orders) : "Not detected"} accent="blue" icon={<CircleDollarSign/>}/>
    </div>
    <div className="chart-grid top">
      <Panel title="Orders & Units Trend (Monthly)" className="wide-panel"><LineTrend data={quarter} profit={[]}/></Panel>
      <Panel title="Regional Order Volume"><BarHorizontal data={region}/></Panel>
      <Panel title="Sales & Avg Ship Duration by Region"><BarHorizontal data={groupAvg(rows,m.region,m.shipping)}/></Panel>
    </div>
    <div className="chart-grid bottom">
      <Panel title="Customer Segment Order Volume"><BarHorizontal data={segment}/></Panel>
      <Panel title="Fulfillment Efficiency: Sales vs Shipping Days" className="wide-panel"><ScatterPlot data={scatter}/></Panel>
      <InsightPanel dataset={dataset} rows={rows} sales={sumBy(rows,m.sales)} profit={sumBy(rows,m.profit)} orders={orders} shipping={shipping}/>
    </div>
  </>;
}

function InsightPanel({dataset,rows,sales,profit,orders,shipping}:{dataset:Dataset;rows:Record<string,unknown>[];sales:number;profit:number;orders:number;shipping?:number}) {
  const m = inferMetrics(dataset);
  const topRegion = groupSum(rows,m.region,m.sales)[0];
  const topCategory = groupSum(rows,m.category,m.sales)[0];
  const topProduct = groupSum(rows,m.product,m.sales)[0];
  return <div className="insight panel"><div className="panel-title">KEY INSIGHTS</div>
    <Insight label="Top Region / State" value={topRegion?.name ?? "—"} accent="cyan"/>
    <Insight label="Top Category" value={topCategory?.name ?? "—"} accent="purple"/>
    <Insight label="Top Product" value={topProduct?.name ?? "—"} accent="green"/>
    <Insight label="Profit / Order" value={orders ? currency(profit/orders) : "—"} accent="blue"/>
    {shipping !== undefined && <Insight label="Average Shipping Days" value={`${shipping.toFixed(2)} Days`} accent="green"/>}
    <div className="insight-foot">Generated from {rows.length.toLocaleString()} filtered rows · {currency(sales)} sales</div>
  </div>;
}

function Insight({label,value,accent}:{label:string;value:string;accent:string}) {
  return <div className={`insight-row ${accent}`}><span>{label}</span><strong>{value}</strong></div>;
}

function LineTrend({data,profit}:{data:{name:string;value:number}[];profit:{name:string;value:number}[]}) {
  const pMap = new Map(profit.map(d=>[d.name,d.value]));
  const merged = data.map(d=>({...d,profit:pMap.get(d.name) ?? null}));
  return <div className="chart"><ResponsiveContainer width="100%" height="100%"><LineChart data={merged}><CartesianGrid stroke="rgba(91,124,190,.16)"/><XAxis dataKey="name" tick={{fill:"#7f91b5",fontSize:10}}/><YAxis tick={{fill:"#7f91b5",fontSize:10}}/><Tooltip contentStyle={{background:"#0b1429",border:"1px solid #20365e",borderRadius:8}}/><Line type="monotone" dataKey="value" name="Sales / Units" stroke="#20d9ff" strokeWidth={3} dot={false}/>{profit.length>0 && <Line type="monotone" dataKey="profit" name="Profit" stroke="#a54dff" strokeWidth={2.5} dot={false}/>}</LineChart></ResponsiveContainer></div>;
}

function formatCompact(value:number) {
  if (Math.abs(value) >= 1000000) return `${(value/1000000).toFixed(1)}M`;
  if (Math.abs(value) >= 1000) return `${(value/1000).toFixed(1)}K`;
  return value.toFixed(0);
}
function formatAxisNumber(value:number) { return formatCompact(value); }

function BarHorizontal({data}:{data:{name:string;value:number}[]}) {
  const display = data.map(d => ({...d, label: d.name.length > 20 ? `${d.name.slice(0,19)}…` : d.name}));
  return <div className="chart"><ResponsiveContainer width="100%" height="100%"><BarChart data={display} layout="vertical" margin={{left:8,right:22}}><CartesianGrid horizontal={false} stroke="rgba(91,124,190,.15)"/><XAxis type="number" tick={{fill:"#7185ad",fontSize:10}} tickFormatter={formatAxisNumber}/><YAxis type="category" dataKey="label" width={90} tick={{fill:"#dbe7ff",fontSize:9}}/><Tooltip cursor={{fill:"rgba(255,255,255,.03)"}} contentStyle={{background:"#0b1429",border:"1px solid #20365e"}} formatter={(v:number)=>[formatCompact(v),"Revenue / Value"]} labelFormatter={(_,payload)=>payload?.[0]?.payload?.name ?? ""}/><Bar dataKey="value" radius={[0,4,4,0]}>{display.map((_,i)=><Cell key={i} fill={chartColors[i%chartColors.length]}/>)}</Bar></BarChart></ResponsiveContainer></div>;
}

function BarVertical({data}:{data:{name:string;value:number}[]}) {
  return <div className="chart"><ResponsiveContainer width="100%" height="100%"><BarChart data={data}><CartesianGrid vertical={false} stroke="rgba(91,124,190,.15)"/><XAxis dataKey="name" tick={{fill:"#dbe7ff",fontSize:9}}/><YAxis tick={{fill:"#7185ad",fontSize:10}}/><Tooltip contentStyle={{background:"#0b1429",border:"1px solid #20365e"}}/><Bar dataKey="value" radius={[3,3,0,0]}>{data.map((_,i)=><Cell key={i} fill={chartColors[i%chartColors.length]}/>)}</Bar></BarChart></ResponsiveContainer></div>;
}

function ScatterPlot({data}:{data:{x:number;y:number;name:string;size:number}[]}) {
  if (!data.length) return <div className="empty-chart"><Sparkles/> Need numeric sales + profit/operations columns to build this view.</div>;
  return <div className="chart"><ResponsiveContainer width="100%" height="100%"><ScatterChart><CartesianGrid stroke="rgba(91,124,190,.15)"/><XAxis type="number" dataKey="x" tick={{fill:"#7185ad",fontSize:10}}/><YAxis type="number" dataKey="y" tick={{fill:"#7185ad",fontSize:10}}/><Tooltip cursor={{strokeDasharray:"3 3"}} contentStyle={{background:"#0b1429",border:"1px solid #20365e"}} formatter={(v:number)=>[v.toLocaleString(),"Value"]}/><Scatter data={data} fill="#8d50ff">{data.map((d,i)=><Cell key={i} r={Math.max(3,Math.min(11,d.size/5))} />)}</Scatter></ScatterChart></ResponsiveContainer></div>;
}

function groupScatter(rows:Record<string,unknown>[], group:string, xcol:string, ycol:string) {
  const map = new Map<string,{x:number;y:number;name:string;size:number}>();
  rows.forEach(r=>{
    const name=String(r[group]??"Unknown");
    const x=Number(String(r[xcol]??"").replace(/[$,%\s,]/g,""));
    const y=Number(String(r[ycol]??"").replace(/[$,%\s,]/g,""));
    if(!Number.isFinite(x)||!Number.isFinite(y)) return;
    const old=map.get(name);
    map.set(name, old ? {...old,x:old.x+x,y:old.y+y,size:old.size+1} : {name,x,y,size:1});
  });
  return [...map.values()].slice(0,40);
}

export default App;
