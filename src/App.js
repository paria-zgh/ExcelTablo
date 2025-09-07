import React, { useState } from "react";
import * as XLSX from "xlsx";
import { saveAs } from "file-saver";
import ExcelJS from "exceljs";
import "./App.css";

function App() {
  const [file1, setFile1] = useState(null);
  const [file2, setFile2] = useState(null);
  const [loading, setLoading] = useState(false);
  const [alertMsg, setAlertMsg] = useState("");

  const handleFile1 = (e) => setFile1(e.target.files[0]);
  const handleFile2 = (e) => setFile2(e.target.files[0]);

  const showAlert = (msg) => {
    setAlertMsg(msg);
    setTimeout(() => setAlertMsg(""), 3000);
  };

  const processFiles = async () => {
    if (!file1 || !file2) {
      showAlert("لطفاً هر دو فایل را انتخاب کنید.");
      return;
    }

    setLoading(true);
    try {
      // --- فایل اول (ترتیب کد عرضه‌ها)
      const data1 = await file1.arrayBuffer();
      const wb1 = XLSX.read(data1, { type: "array" });
      const ws1 = wb1.Sheets[wb1.SheetNames[0]];
      const df1 = XLSX.utils.sheet_to_json(ws1);
      const order = [...new Set(df1.map((row) => String(row["کد عرضه"])))];

      // --- فایل دوم (داده‌ها)
      const data2 = await file2.arrayBuffer();
      const wb2 = XLSX.read(data2, { type: "array" });
      const ws2 = wb2.Sheets[wb2.SheetNames[0]];
      const df2 = XLSX.utils.sheet_to_json(ws2);

      // --- ستون‌های خروجی (به ترتیب از راست به چپ)
      const keepColumns = [
        "عرضه",
        "تقاضا",
        "نام کالا",
        "نام مشتری",
        "محموله",
        "نام عرضه کننده",
        "قيمت پايه عرضه",
        "کد عرضه",
      ];
      const availableCols = keepColumns.filter((c) => c in (df2[0] || {}));

      // --- مرتب‌سازی بر اساس کد عرضه فایل اول
      let result = [];
      order.forEach((code) => {
        const subset = df2
          .filter((row) => String(row["کد عرضه"]) === code)
          .map((row) => {
            let filtered = {};
            availableCols.forEach((col) => {
              filtered[col] = row[col] || "";
            });
            return filtered;
          });
        if (subset.length > 0) {
          result.push(...subset);
          result.push({}); // ردیف خالی
        }
      });

      // --- ساخت Workbook با ExcelJS
      const workbook = new ExcelJS.Workbook();
      const sheet = workbook.addWorksheet("نتیجه مرتب‌سازی");

      // راست به چپ کردن کل شیت
      sheet.views = [{ rightToLeft: true }];

      // اضافه کردن هدر
      sheet.addRow(availableCols);

      // اضافه کردن داده‌ها
      result.forEach((row) => {
        sheet.addRow(availableCols.map((c) => row[c] || ""));
      });

      // استایل‌دهی همه سلول‌ها
      sheet.eachRow((row) => {
        row.eachCell((cell) => {
          cell.font = { name: "B Nazanin", size: 12 };
          cell.alignment = { vertical: "middle", horizontal: "center" };
          cell.border = {
            top: { style: "thin" },
            left: { style: "thin" },
            bottom: { style: "thin" },
            right: { style: "thin" },
          };
          // اگر سلول عدد بود → فرمت هزارگان
          if (typeof cell.value === "number") {
            cell.numFmt = "#,##0";
          }
          // پس‌زمینه سفید
          cell.fill = {
            type: "pattern",
            pattern: "solid",
            fgColor: { argb: "FFFFFFFF" },
          };
        });
      });

      // خروجی فایل
      const buffer = await workbook.xlsx.writeBuffer();
      saveAs(new Blob([buffer]), "اکسل_مرتب.xlsx");
    } catch (error) {
      showAlert("خطا در پردازش فایل‌ها!");
      console.error(error);
    } finally {
      setLoading(false);
    }
  };

  return (
    <div className="container">
      <h2>📊 مرتب‌سازی اکسل دوم بر اساس اکسل اول</h2>

      <div className="file-input">
        <label htmlFor="file1" className="custom-file-btn">
          انتخاب فایل اول
        </label>
        <input id="file1" type="file" accept=".xlsx,.xls" onChange={handleFile1} />
        <p className="file-name">{file1 ? file1.name : "هیچ فایلی انتخاب نشده"}</p>
      </div>

      <div className="file-input">
        <label htmlFor="file2" className="custom-file-btn">
          انتخاب فایل دوم
        </label>
        <input id="file2" type="file" accept=".xlsx,.xls" onChange={handleFile2} />
        <p className="file-name">{file2 ? file2.name : "هیچ فایلی انتخاب نشده"}</p>
      </div>

      <button onClick={processFiles} disabled={loading || !file1 || !file2}>
        {loading ? "در حال پردازش..." : "دانلود اکسل مرتب‌شده"}
      </button>

      {alertMsg && <div className="custom-alert">{alertMsg}</div>}
    </div>
  );
}

export default App;
