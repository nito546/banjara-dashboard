import { useState, useEffect, useRef } from 'react';
import { BarChart, Bar, XAxis, YAxis, Tooltip, ResponsiveContainer, PieChart, Pie, Cell, Legend } from 'recharts';
import * as XLSX from 'xlsx';
import './App.css';

const ALL_ORDERS = [
  {name:"#1109",date:"2026-05-20",total:6599},
  {name:"#1108",date:"2026-05-20",total:4500},
  {name:"#1107",date:"2026-05-20",total:4900},
  {name:"#1106",date:"2026-05-19",total:4455},
  {name:"#1105",date:"2026-05-18",total:4900},
  {name:"#1104",date:"2026-05-18",total:4455},
  {name:"#1103",date:"2026-05-18",total:4900},
  {name:"#1102",date:"2026-05-17",total:14345},
  {name:"#1101",date:"2026-05-17",total:4500},
  {name:"#1100",date:"2026-05-14",total:4500},
  {name:"#1099",date:"2026-05-14",total:4900},
  {name:"#1098",date:"2026-05-13",total:4900},
  {name:"#1097",date:"2026-05-13",total:4500},
  {name:"#1096",date:"2026-05-12",total:4500},
  {name:"#1095",date:"2026-05-12",total:4500},
  {name:"#1094",date:"2026-05-12",total:4500},
  {name:"#1093",date:"2026-05-12",total:4500},
  {name:"#1092",date:"2026-05-11",total:4900},
  {name:"#1091",date:"2026-05-11",total:4900},
  {name:"#1090",date:"2026-05-11",total:4500},
  {name:"#1089",date:"2026-05-10",total:4900},
  {name:"#1088",date:"2026-05-09",total:4500},
  {name:"#1087",date:"2026-05-09",total:4900},
  {name:"#1086",date:"2026-05-09",total:21600},
  {name:"#1085",date:"2026-05-09",total:4900},
  {name:"#1084",date:"2026-05-08",total:1250},
  {name:"#1083",date:"2026-05-08",total:4900},
  {name:"#1082",date:"2026-05-08",total:4900},
  {name:"#1081",date:"2026-05-07",total:4900},
  {name:"#1080",date:"2026-05-06",total:4500},
  {name:"#1079",date:"2026-05-06",total:4500},
  {name:"#1078",date:"2026-05-05",total:4900},
  {name:"#1077",date:"2026-05-05",total:4900},
  {name:"#1076",date:"2026-05-05",total:4500},
  {name:"#1075",date:"2026-05-05",total:4900},
  {name:"#1074",date:"2026-05-02",total:4900},
  {name:"#1073",date:"2026-05-01",total:4900},
  {name:"#1072",date:"2026-05-01",total:5550},
  {name:"#1071",date:"2026-05-01",total:4900},
  {name:"#1070",date:"2026-04-30",total:4500},
  {name:"#1069",date:"2026-04-30",total:17500},
  {name:"#1068",date:"2026-04-30",total:4500},
  {name:"#1067",date:"2026-04-29",total:6650},
  {name:"#1066",date:"2026-04-29",total:4500},
  {name:"#1065",date:"2026-04-29",total:4900},
  {name:"#1064",date:"2026-04-27",total:4900},
  {name:"#1063",date:"2026-04-26",total:4900},
  {name:"#1062",date:"2026-04-26",total:4900},
  {name:"#1061",date:"2026-04-25",total:4500},
  {name:"#1060",date:"2026-04-25",total:4900},
  {name:"#1059",date:"2026-04-24",total:4500},
  {name:"#1058",date:"2026-04-24",total:4500},
  {name:"#1057",date:"2026-04-24",total:4900},
  {name:"#1056",date:"2026-04-24",total:4500},
  {name:"#1055",date:"2026-04-24",total:4900},
  {name:"#1065b",date:"2026-04-20",total:4050},
  {name:"#1064b",date:"2026-04-20",total:4500},
  {name:"#1063b",date:"2026-04-20",total:4500},
  {name:"#1062b",date:"2026-04-20",total:4500},
  {name:"#1060b",date:"2026-04-20",total:4500},
  {name:"#1061b",date:"2026-04-20",total:4050},
  {name:"#1058b",date:"2026-04-20",total:4500},
  {name:"#1057b",date:"2026-04-20",total:4050},
  {name:"#1056b",date:"2026-04-20",total:4900},
  {name:"#1055b",date:"2026-04-20",total:4900},
  {name:"#1054",date:"2026-04-20",total:4500},
  {name:"#1053",date:"2026-04-20",total:4500},
  {name:"#1052",date:"2026-04-20",total:4900},
  {name:"#1051",date:"2026-04-20",total:4900},
  {name:"#1050",date:"2026-04-20",total:4900},
  {name:"#1049",date:"2026-04-20",total:4900},
  {name:"#1047",date:"2026-04-20",total:22500},
  {name:"#1046",date:"2026-04-20",total:4500},
  {name:"#1045",date:"2026-04-20",total:4500},
  {name:"#1044",date:"2026-04-20",total:4500},
  {name:"#1043",date:"2026-04-20",total:4500},
  {name:"#1042",date:"2026-04-20",total:4500},
  {name:"#1041",date:"2026-04-20",total:45},
  {name:"#1040",date:"2026-04-20",total:4050},
  {name:"#1039",date:"2026-04-20",total:4500},
  {name:"#1038",date:"2026-04-20",total:4500},
  {name:"#1033",date:"2026-04-20",total:11550},
  {name:"#1032",date:"2026-04-20",total:4500},
  {name:"#1031",date:"2026-04-20",total:4050},
  {name:"#1030",date:"2026-04-20",total:4050},
  {name:"#1029",date:"2026-04-20",total:4500},
  {name:"#1028",date:"2026-04-20",total:4500},
  {name:"#1027",date:"2026-04-20",total:4050},
  {name:"#1026",date:"2026-04-20",total:4500},
  {name:"#1025",date:"2026-04-20",total:4500},
  {name:"#1024",date:"2026-04-20",total:4050},
  {name:"#1023",date:"2026-04-20",total:30900},
  {name:"#1022",date:"2026-04-20",total:4500},
  {name:"#1021",date:"2026-04-20",total:4500},
  {name:"#1020",date:"2026-04-20",total:4500},
  {name:"#1019",date:"2026-04-20",total:15900},
  {name:"#1018",date:"2026-04-17",total:60},
  {name:"#1017",date:"2026-04-17",total:27},
];

const CAT_COLORS = {
  Shipping: '#378ADD',
  Labour: '#1D9E75',
  Material: '#D85A30',
  Utilities: '#BA7517',
  Ads: '#D4537E',
  Other: '#888780',
};

const fmt = (n) => 'PKR ' + Math.round(n).toLocaleString();

export default function App() {
  const today = new Date().toISOString().slice(0, 10);
  const firstOrder = '2026-04-17';

  const [startDate, setStartDate] = useState(firstOrder);
  const [endDate, setEndDate] = useState(today);
  const [expenses, setExpenses] = useState(() => {
    try { return JSON.parse(localStorage.getItem('banjara_expenses') || '[]'); } catch { return []; }
  });
  const [inpCat, setInpCat] = useState('Shipping');
  const [inpDesc, setInpDesc] = useState('');
  const [inpDate, setInpDate] = useState(today);
  const [inpAmt, setInpAmt] = useState('');
  const fileRef = useRef();

  useEffect(() => {
    try { localStorage.setItem('banjara_expenses', JSON.stringify(expenses)); } catch {}
  }, [expenses]);

  const filteredOrders = ALL_ORDERS.filter(o => o.date >= startDate && o.date <= endDate);
  const filteredExpenses = expenses.filter(e => e.date >= startDate && e.date <= endDate);

  const totalRevenue = filteredOrders.reduce((s, o) => s + o.total, 0);
  const totalExpenses = filteredExpenses.reduce((s, e) => s + e.amount, 0);
  const netProfit = totalRevenue - totalExpenses;
  const margin = totalRevenue > 0 ? (netProfit / totalRevenue * 100) : 0;
  const avgOrder = filteredOrders.length > 0 ? totalRevenue / filteredOrders.length : 0;

  // Weekly revenue chart data
  const weeks = {};
  filteredOrders.forEach(o => {
    const d = new Date(o.date);
    const mon = new Date(d);
    mon.setDate(d.getDate() - ((d.getDay() + 6) % 7));
    const key = mon.toISOString().slice(0, 10);
    weeks[key] = (weeks[key] || 0) + o.total;
  });
  const weeklyData = Object.keys(weeks).sort().map(k => ({
    week: k.slice(5),
    revenue: Math.round(weeks[k]),
  }));

  // Expense by category
  const catTotals = {};
  filteredExpenses.forEach(e => { catTotals[e.cat] = (catTotals[e.cat] || 0) + e.amount; });
  const pieData = Object.keys(catTotals).map(c => ({ name: c, value: Math.round(catTotals[c]) }));

  const plData = [
    { name: 'Revenue', amount: Math.round(totalRevenue) },
    { name: 'Expenses', amount: Math.round(totalExpenses) },
    { name: 'Net Profit', amount: Math.round(netProfit) },
  ];

  function addExpense() {
    if (!inpDesc || !inpDate || !inpAmt || isNaN(parseFloat(inpAmt))) return alert('Fill in all fields');
    setExpenses(prev => [...prev, { cat: inpCat, desc: inpDesc, date: inpDate, amount: parseFloat(inpAmt) }]);
    setInpDesc('');
    setInpAmt('');
  }

  function removeExpense(i) {
    setExpenses(prev => prev.filter((_, idx) => idx !== i));
  }

  function handleExcelUpload(e) {
    const file = e.target.files[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = (evt) => {
      const wb = XLSX.read(evt.target.result, { type: 'binary' });
      const ws = wb.Sheets[wb.SheetNames[0]];
      const rows = XLSX.utils.sheet_to_json(ws);
      const parsed = rows.map(r => ({
        cat: r.Category || r.category || 'Other',
        desc: r.Description || r.description || '',
        date: r.Date || r.date || today,
        amount: parseFloat(r.Amount || r.amount || 0),
      })).filter(r => r.amount > 0);
      setExpenses(prev => [...prev, ...parsed]);
    };
    reader.readAsBinaryString(file);
    e.target.value = '';
  }

  function downloadTemplate() {
    const ws = XLSX.utils.aoa_to_sheet([
      ['Date', 'Category', 'Description', 'Amount'],
      ['2026-05-01', 'Shipping', 'TCS courier batch', 8000],
      ['2026-05-03', 'Labour', 'Staff wages', 25000],
      ['2026-05-10', 'Ads', 'Meta ads', 15000],
      ['2026-05-15', 'Material', 'Raw materials', 12000],
      ['2026-05-20', 'Utilities', 'Electricity bill', 3500],
    ]);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Expenses');
    XLSX.writeFile(wb, 'banjara-expenses-template.xlsx');
  }

  const setPreset = (preset) => {
    const now = new Date();
    if (preset === 'thisMonth') {
      setStartDate(new Date(now.getFullYear(), now.getMonth(), 1).toISOString().slice(0, 10));
      setEndDate(today);
    } else if (preset === 'lastMonth') {
      const f = new Date(now.getFullYear(), now.getMonth() - 1, 1);
      const l = new Date(now.getFullYear(), now.getMonth(), 0);
      setStartDate(f.toISOString().slice(0, 10));
      setEndDate(l.toISOString().slice(0, 10));
    } else if (preset === 'last30') {
      const f = new Date(now); f.setDate(now.getDate() - 30);
      setStartDate(f.toISOString().slice(0, 10));
      setEndDate(today);
    } else if (preset === 'allTime') {
      setStartDate(firstOrder);
      setEndDate(today);
    }
  };

  return (
    <div style={{ maxWidth: 900, margin: '0 auto', padding: '24px 16px', fontFamily: 'system-ui, sans-serif', color: '#1a1a1a' }}>

      {/* Header */}
      <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: 24 }}>
        <div>
          <h1 style={{ fontSize: 22, fontWeight: 600, margin: 0 }}>Banjara Gear</h1>
          <p style={{ fontSize: 13, color: '#888', margin: '2px 0 0' }}>Profit & Loss Dashboard</p>
        </div>
        <div style={{ fontSize: 12, color: '#888', background: '#f5f5f5', padding: '6px 12px', borderRadius: 8 }}>
          banjaragear.com
        </div>
      </div>

      {/* Date Filter */}
      <div style={{ background: '#f9f9f9', border: '1px solid #eee', borderRadius: 12, padding: '16px', marginBottom: 24 }}>
        <div style={{ fontSize: 11, fontWeight: 500, color: '#888', textTransform: 'uppercase', letterSpacing: '0.06em', marginBottom: 12 }}>Date range</div>
        <div style={{ display: 'flex', flexWrap: 'wrap', gap: 10, alignItems: 'center' }}>
          <input type="date" value={startDate} onChange={e => setStartDate(e.target.value)}
            style={{ padding: '6px 10px', border: '1px solid #ddd', borderRadius: 8, fontSize: 13 }} />
          <span style={{ color: '#888', fontSize: 13 }}>to</span>
          <input type="date" value={endDate} onChange={e => setEndDate(e.target.value)}
            style={{ padding: '6px 10px', border: '1px solid #ddd', borderRadius: 8, fontSize: 13 }} />
          <div style={{ display: 'flex', gap: 6, flexWrap: 'wrap' }}>
            {[['thisMonth','This month'],['lastMonth','Last month'],['last30','Last 30 days'],['allTime','All time']].map(([k,l]) => (
              <button key={k} onClick={() => setPreset(k)}
                style={{ fontSize: 12, padding: '5px 10px', border: '1px solid #ddd', borderRadius: 8, background: '#fff', cursor: 'pointer' }}>
                {l}
              </button>
            ))}
          </div>
        </div>
      </div>

      {/* Revenue Metrics */}
      <div style={{ fontSize: 11, fontWeight: 500, color: '#888', textTransform: 'uppercase', letterSpacing: '0.06em', marginBottom: 10 }}>Revenue</div>
      <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fit, minmax(140px, 1fr))', gap: 10, marginBottom: 24 }}>
        {[
          { label: 'Total revenue', value: fmt(totalRevenue) },
          { label: 'Orders', value: filteredOrders.length },
          { label: 'Avg order', value: fmt(avgOrder) },
          { label: 'Highest order', value: filteredOrders.length ? fmt(Math.max(...filteredOrders.map(o => o.total))) : 'PKR 0' },
        ].map(m => (
          <div key={m.label} style={{ background: '#f5f5f5', borderRadius: 10, padding: '14px 16px' }}>
            <div style={{ fontSize: 12, color: '#888', marginBottom: 4 }}>{m.label}</div>
            <div style={{ fontSize: 20, fontWeight: 600 }}>{m.value}</div>
          </div>
        ))}
      </div>

      {/* Charts Row */}
      <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: 16, marginBottom: 24 }}>
        <div style={{ background: '#fff', border: '1px solid #eee', borderRadius: 12, padding: 16 }}>
          <div style={{ fontSize: 11, fontWeight: 500, color: '#888', textTransform: 'uppercase', letterSpacing: '0.06em', marginBottom: 12 }}>Weekly revenue</div>
          <ResponsiveContainer width="100%" height={180}>
            <BarChart data={weeklyData}>
              <XAxis dataKey="week" tick={{ fontSize: 10 }} />
              <YAxis tick={{ fontSize: 10 }} tickFormatter={v => 'â‚¨' + Math.round(v / 1000) + 'k'} />
              <Tooltip formatter={v => ['PKR ' + v.toLocaleString(), 'Revenue']} />
              <Bar dataKey="revenue" fill="#378ADD" radius={[4, 4, 0, 0]} />
            </BarChart>
          </ResponsiveContainer>
        </div>

        <div style={{ background: '#fff', border: '1px solid #eee', borderRadius: 12, padding: 16 }}>
          <div style={{ fontSize: 11, fontWeight: 500, color: '#888', textTransform: 'uppercase', letterSpacing: '0.06em', marginBottom: 12 }}>Expense breakdown</div>
          {pieData.length === 0 ? (
            <div style={{ height: 180, display: 'flex', alignItems: 'center', justifyContent: 'center', color: '#bbb', fontSize: 13 }}>No expenses yet</div>
          ) : (
            <ResponsiveContainer width="100%" height={180}>
              <PieChart>
                <Pie data={pieData} dataKey="value" nameKey="name" cx="50%" cy="50%" outerRadius={65} label={({ name, percent }) => `${name} ${(percent * 100).toFixed(0)}%`} labelLine={false} fontSize={11}>
                  {pieData.map((entry, i) => <Cell key={i} fill={CAT_COLORS[entry.name] || '#888'} />)}
                </Pie>
                <Tooltip formatter={v => ['PKR ' + v.toLocaleString()]} />
              </PieChart>
            </ResponsiveContainer>
          )}
        </div>
      </div>

      {/* Expenses Table */}
      <div style={{ background: '#fff', border: '1px solid #eee', borderRadius: 12, padding: 16, marginBottom: 24 }}>
        <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: 16 }}>
          <span style={{ fontSize: 15, fontWeight: 600 }}>Expenses</span>
          <div style={{ display: 'flex', gap: 8 }}>
            <button onClick={downloadTemplate}
              style={{ fontSize: 12, padding: '5px 12px', border: '1px solid #ddd', borderRadius: 8, background: '#fff', cursor: 'pointer' }}>
              Download template
            </button>
            <button onClick={() => fileRef.current.click()}
              style={{ fontSize: 12, padding: '5px 12px', border: '1px solid #378ADD', borderRadius: 8, background: '#EBF4FF', color: '#378ADD', cursor: 'pointer' }}>
              Upload Excel
            </button>
            <input ref={fileRef} type="file" accept=".xlsx,.xls" style={{ display: 'none' }} onChange={handleExcelUpload} />
          </div>
        </div>

        <table style={{ width: '100%', fontSize: 13, borderCollapse: 'collapse' }}>
          <thead>
            <tr style={{ borderBottom: '1px solid #eee' }}>
              {['Category','Description','Date','Amount (PKR)',''].map(h => (
                <th key={h} style={{ textAlign: h === 'Amount (PKR)' ? 'right' : 'left', padding: '6px 8px', fontSize: 11, color: '#888', fontWeight: 500 }}>{h}</th>
              ))}
            </tr>
          </thead>
          <tbody>
            {filteredExpenses.length === 0 ? (
              <tr><td colSpan={5} style={{ padding: '20px 8px', color: '#bbb', fontSize: 13 }}>No expenses in this period â€” add one below or upload an Excel file</td></tr>
            ) : filteredExpenses.map((e, i) => (
              <tr key={i} style={{ borderBottom: '1px solid #f5f5f5' }}>
                <td style={{ padding: '8px' }}>
                  <span style={{ background: CAT_COLORS[e.cat] + '22', color: CAT_COLORS[e.cat], fontSize: 11, padding: '2px 8px', borderRadius: 999 }}>{e.cat}</span>
                </td>
                <td style={{ padding: '8px' }}>{e.desc}</td>
                <td style={{ padding: '8px', color: '#888' }}>{e.date}</td>
                <td style={{ padding: '8px', textAlign: 'right', fontWeight: 500 }}>{Math.round(e.amount).toLocaleString()}</td>
                <td style={{ padding: '8px' }}>
                  <button onClick={() => removeExpense(expenses.indexOf(e))}
                    style={{ fontSize: 12, padding: '2px 8px', border: '1px solid #fcc', borderRadius: 6, background: '#fff', color: '#e24b4a', cursor: 'pointer' }}>Ã—</button>
                </td>
              </tr>
            ))}
          </tbody>
        </table>

        {/* Add Expense Row */}
        <div style={{ display: 'flex', gap: 8, marginTop: 16, flexWrap: 'wrap' }}>
          <select value={inpCat} onChange={e => setInpCat(e.target.value)}
            style={{ padding: '7px 10px', border: '1px solid #ddd', borderRadius: 8, fontSize: 13 }}>
            {Object.keys(CAT_COLORS).map(c => <option key={c}>{c}</option>)}
          </select>
          <input value={inpDesc} onChange={e => setInpDesc(e.target.value)} placeholder="Description"
            style={{ flex: 1, minWidth: 120, padding: '7px 10px', border: '1px solid #ddd', borderRadius: 8, fontSize: 13 }} />
          <input type="date" value={inpDate} onChange={e => setInpDate(e.target.value)}
            style={{ padding: '7px 10px', border: '1px solid #ddd', borderRadius: 8, fontSize: 13 }} />
          <input type="number" value={inpAmt} onChange={e => setInpAmt(e.target.value)} placeholder="Amount"
            style={{ width: 120, padding: '7px 10px', border: '1px solid #ddd', borderRadius: 8, fontSize: 13 }} />
          <button onClick={addExpense}
            style={{ padding: '7px 16px', background: '#1a1a1a', color: '#fff', border: 'none', borderRadius: 8, fontSize: 13, cursor: 'pointer' }}>
            + Add
          </button>
        </div>
      </div>

      {/* P&L Summary */}
      <div style={{ fontSize: 11, fontWeight: 500, color: '#888', textTransform: 'uppercase', letterSpacing: '0.06em', marginBottom: 10 }}>Profit & loss summary</div>
      <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fit, minmax(140px, 1fr))', gap: 10, marginBottom: 24 }}>
        <div style={{ background: '#f0faf5', borderRadius: 10, padding: '14px 16px' }}>
          <div style={{ fontSize: 12, color: '#888', marginBottom: 4 }}>Total revenue</div>
          <div style={{ fontSize: 20, fontWeight: 600, color: '#1D9E75' }}>{fmt(totalRevenue)}</div>
        </div>
        <div style={{ background: '#fff5f5', borderRadius: 10, padding: '14px 16px' }}>
          <div style={{ fontSize: 12, color: '#888', marginBottom: 4 }}>Total expenses</div>
          <div style={{ fontSize: 20, fontWeight: 600, color: '#e24b4a' }}>{fmt(totalExpenses)}</div>
        </div>
        <div style={{ background: netProfit >= 0 ? '#f0faf5' : '#fff5f5', borderRadius: 10, padding: '14px 16px' }}>
          <div style={{ fontSize: 12, color: '#888', marginBottom: 4 }}>Net profit</div>
          <div style={{ fontSize: 20, fontWeight: 600, color: netProfit >= 0 ? '#1D9E75' : '#e24b4a' }}>{fmt(netProfit)}</div>
        </div>
        <div style={{ background: '#f5f5f5', borderRadius: 10, padding: '14px 16px' }}>
          <div style={{ fontSize: 12, color: '#888', marginBottom: 4 }}>Profit margin</div>
          <div style={{ fontSize: 20, fontWeight: 600, color: margin >= 0 ? '#1D9E75' : '#e24b4a' }}>{margin.toFixed(1)}%</div>
        </div>
      </div>

      {/* P&L Bar Chart */}
      <div style={{ background: '#fff', border: '1px solid #eee', borderRadius: 12, padding: 16 }}>
        <div style={{ fontSize: 11, fontWeight: 500, color: '#888', textTransform: 'uppercase', letterSpacing: '0.06em', marginBottom: 12 }}>Revenue vs expenses</div>
        <ResponsiveContainer width="100%" height={200}>
          <BarChart data={plData}>
            <XAxis dataKey="name" tick={{ fontSize: 12 }} />
            <YAxis tick={{ fontSize: 10 }} tickFormatter={v => 'â‚¨' + Math.round(v / 1000) + 'k'} />
            <Tooltip formatter={v => ['PKR ' + Math.round(v).toLocaleString()]} />
            <Bar dataKey="amount" radius={[4, 4, 0, 0]}>
              {plData.map((entry, i) => (
                <Cell key={i} fill={i === 0 ? '#1D9E75' : i === 1 ? '#e24b4a' : entry.amount >= 0 ? '#378ADD' : '#e24b4a'} />
              ))}
            </Bar>
          </BarChart>
        </ResponsiveContainer>
      </div>

      <div style={{ textAlign: 'center', fontSize: 12, color: '#ccc', marginTop: 24 }}>Banjara Gear Dashboard â€” banjaragear.com</div>
    </div>
  );
}
