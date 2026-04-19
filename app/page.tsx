"use client";
import { useState } from "react";

export default function Home() {
  const [startDate, setStartDate] = useState("២១");
  const [endDate, setEndDate] = useState("០១ ខែមីនា ឆ្នាំ២០២៦");
  const [reportDate, setReportDate] = useState("ត្រូវនឹងថ្ងៃទី០១ ខែមីនា ឆ្នាំ២០២៦");
  const [isGenerating, setIsGenerating] = useState(false);

  const [rows, setRows] = useState([
    ["ប៉ុស្តិ៍នគរបាលរដ្ឋបាលក្រាំងចេក", 5, 2, 0, 0, 2, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0],
  ]);

  const addRow = () => {
    setRows([...rows, ["គោលដៅថ្មី", 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0]]);
  };

  const updateCell = (rowIndex: number, colIndex: number, value: string) => {
    const newRows = [...rows];
    newRows[rowIndex][colIndex] = colIndex === 0 ? value : parseInt(value) || 0;
    setRows(newRows);
  };

  const removeRow = (index: number) => {
    setRows(rows.filter((_, i) => i !== index));
  };

  const downloadExcel = async () => {
    setIsGenerating(true);
    try {
      const response = await fetch("/api/generate", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ startDate, endDate, reportDate, rows }),
      });
      const blob = await response.blob();
      const url = window.URL.createObjectURL(blob);
      const a = document.createElement("a");
      a.href = url;
      a.download = "Police_Report.xlsx";
      a.click();
    } catch (error) {
      alert("Error generating Excel");
    }
    setIsGenerating(false);
  };

  const downloadPDF = () => {
    let totals = Array(15).fill(0);
    let rowsHtml = rows.map((row, idx) => {
      let tr = `<tr><td>${idx + 1}</td><td style='text-align:left;'>${row[0]}</td>`;
      for (let c = 1; c < 16; c++) {
        let val = Number(row[c]);
        tr += `<td>${val > 0 ? val : ""}</td>`;
        totals[c - 1] += val;
      }
      return tr + "</tr>";
    }).join("");

    let totalsHtml = "<tr><td colspan='2' style='font-weight:bold;'>សរុប</td>";
    totals.forEach(t => (totalsHtml += `<td style='font-weight:bold;'>${t > 0 ? t : ""}</td>`));
    totalsHtml += "</tr>";

    const htmlContent = `
      <!DOCTYPE html>
      <html lang="km">
      <head>
          <meta charset="UTF-8">
          <title>Police Report</title>
          <style>
              body { font-family: 'Khmer OS Battambang', Arial, sans-serif; font-size: 14px; padding: 20px; }
              .moul { font-family: 'Khmer OS Moul Light', cursive; }
              .center { text-align: center; }
              table { width: 100%; border-collapse: collapse; margin-top: 20px; }
              th, td { border: 1px solid black; padding: 5px; text-align: center; }
              .right-text { text-align: right; padding-right: 50px; }
          </style>
      </head>
      <body>
          <div style="float: right; text-align: center;" class="moul">
              ព្រះរាជាណាចក្រកម្ពុជា<br>ជាតិ សាសនា ព្រះមហាក្សត្រ<br>
          </div>
          <div style="clear: both;"></div>
          <div class="moul" style="margin-top: -30px;">អធិការដ្ឋាននគរបាលស្រុកសាមគ្គីមុនីជ័យ</div>
          
          <h3 class="moul center" style="margin-top: 40px;">លទ្ធផលការងារសង្គមរបស់កងកម្លាំង</h3>
          <div class="center">ប្រចាំសប្តាហ៍ គិតពីថ្ងៃទី ${startDate} ដល់ ថ្ងៃទី​ ${endDate}</div>

          <table>
              <tr style="font-weight: bold; background-color: #f2f2f2;">
                  <td rowspan="3">ល.រ</td><td rowspan="3">គោលដៅចុះធ្វើការ</td><td colspan="2">សម្រួលចរាចរណ៍</td>
                  <td colspan="4">ការផ្តល់សេវារដ្ឋបាល</td><td colspan="5">អន្តរគមន៍ជួយសង្គ្រោះប្រជាពលរដ្ឋ</td>
                  <td colspan="4">ការចុះជួយប្រ/រដ្ឋ</td>
              </tr>
              <tr style="font-weight: bold; background-color: #f2f2f2;">
                  <td rowspan="2">សាលារៀន</td><td rowspan="2">ទីប្រជុំជន</td>
                  <td rowspan="2">អត្ត.<br>ខ្មែរ</td><td rowspan="2">សៀវភៅ<br>ស្នាក់នៅ</td><td rowspan="2">សៀវភៅ<br>គ្រួសារ</td><td rowspan="2">ផ្សេងៗ</td>
                  <td colspan="3">គ្រោះថ្នាក់</td><td rowspan="2">អ្នកឆ្លង<br>ទន្លេ</td><td rowspan="2">ផ្សេងៗ</td>
                  <td rowspan="2">ជួសជុល<br>ផ្ទះ</td><td rowspan="2">ជួយប្រ<br>មូលផល</td><td rowspan="2">ជួយជន<br>ក្រីក្រ</td><td rowspan="2">ផ្សេងៗ</td>
              </tr>
              <tr style="font-weight: bold; background-color: #f2f2f2;">
                  <td>ចរាចរណ៍</td><td>អគ្គីភ័យ</td><td>អាវុធ.<br>ជាតិផ្ទុះ</td>
              </tr>
              ${rowsHtml}
              ${totalsHtml}
          </table>
          <div class="right-text" style="margin-top: 30px;">
              ថ្ងៃអាទិត្យ ១៣កើត ខែផល្គុន ឆ្នាំម្សាញ់ សប្តស័ក ពុទ្ធសករាជ ២៥៦៩<br>
              ${reportDate}<br><br>
              <b class="moul">នាយប៉ុស្តិ៍នគរបាលរដ្ឋបាលក្រាំងចេក</b>
          </div>
          <script>window.onload = function() { window.print(); }</script>
      </body>
      </html>
    `;
    const blob = new Blob([htmlContent], { type: "text/html" });
    window.open(URL.createObjectURL(blob), "_blank");
  };

  const columns = ["គោលដៅ", "សាលារៀន", "ទីប្រជុំជន", "អត្ត.ខ្មែរ", "ស្នាក់នៅ", "គ្រួសារ", "សេវាផ្សេងៗ", "គ្រោះថ្នាក់ចរាចរណ៍", "អគ្គីភ័យ", "អាវុធផ្ទុះ", "ឆ្លងទន្លេ", "សង្គ្រោះផ្សេងៗ", "ផ្ទះ", "មូលផល", "ជនក្រីក្រ", "ជួយផ្សេងៗ"];

  return (
    <main className="p-8 font-sans max-w-7xl mx-auto">
      <h1 className="text-3xl font-bold mb-2">ប្រព័ន្ធបញ្ចូលទិន្នន័យរបាយការណ៍កងកម្លាំង</h1>
      <p className="text-gray-600 mb-8">បញ្ចូលទិន្នន័យនៅទីនេះ រួចទាញយកជា Excel ឬ PDF (Vercel Ready)</p>

      <div className="grid grid-cols-1 md:grid-cols-3 gap-4 mb-8">
        <div>
          <label className="block text-sm font-medium mb-1">ថ្ងៃទីចាប់ផ្តើម</label>
          <input type="text" className="w-full border p-2 rounded" value={startDate} onChange={(e) => setStartDate(e.target.value)} />
        </div>
        <div>
          <label className="block text-sm font-medium mb-1">ថ្ងៃទីបញ្ចប់</label>
          <input type="text" className="w-full border p-2 rounded" value={endDate} onChange={(e) => setEndDate(e.target.value)} />
        </div>
        <div>
          <label className="block text-sm font-medium mb-1">ថ្ងៃធ្វើរបាយការណ៍</label>
          <input type="text" className="w-full border p-2 rounded" value={reportDate} onChange={(e) => setReportDate(e.target.value)} />
        </div>
      </div>

      <div className="overflow-x-auto mb-4 border rounded shadow">
        <table className="w-full text-sm text-left">
          <thead className="bg-gray-100 uppercase text-xs">
            <tr>
              {columns.map((col, i) => <th key={i} className="px-4 py-3 border">{col}</th>)}
              <th className="px-4 py-3 border">X</th>
            </tr>
          </thead>
          <tbody>
            {rows.map((row, rIdx) => (
              <tr key={rIdx} className="border-b">
                {row.map((cell, cIdx) => (
                  <td key={cIdx} className="border p-0">
                    <input
                      type={cIdx === 0 ? "text" : "number"}
                      className="w-full h-full p-2 outline-none"
                      value={cell === 0 ? "" : cell}
                      placeholder="0"
                      onChange={(e) => updateCell(rIdx, cIdx, e.target.value)}
                    />
                  </td>
                ))}
                <td className="border text-center">
                  <button onClick={() => removeRow(rIdx)} className="text-red-500 font-bold px-2 hover:bg-red-100 rounded">X</button>
                </td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>

      <button onClick={addRow} className="bg-green-600 text-white px-4 py-2 rounded shadow hover:bg-green-700 mb-8">+ បន្ថែមគោលដៅ</button>

      <div className="flex gap-4 p-4 bg-gray-50 rounded-lg shadow-inner">
        <button onClick={downloadExcel} disabled={isGenerating} className="flex-1 bg-blue-600 text-white py-3 rounded-lg font-bold shadow hover:bg-blue-700 transition">
          {isGenerating ? "កំពុងបង្កើត Excel..." : "📥 ទាញយកជា Excel"}
        </button>
        <button onClick={downloadPDF} className="flex-1 bg-red-600 text-white py-3 rounded-lg font-bold shadow hover:bg-red-700 transition">
          📄 ទាញយកជា PDF (Print)
        </button>
      </div>
    </main>
  );
}
