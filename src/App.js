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

  const isNumeric = (v) => {
    if (v === null || v === undefined || v === "") return false;
    // remove possible thousand separators and test
    const s = String(v).toString().replace(/,/g, "").trim();
    return !isNaN(parseFloat(s)) && isFinite(s);
  };

  const processFiles = async () => {
    if (!file1 || !file2) {
      showAlert("لطفاً هر دو فایل را انتخاب کنید.");
      return;
    }

    setLoading(true);
    try {
      // --- فایل اول (ترتیب کد عرضه‌ها + ممکن است حاوی مقدار عرضه/تقاضا باشد)
      const data1 = await file1.arrayBuffer();
      const wb1 = XLSX.read(data1, { type: "array" });
      const ws1 = wb1.Sheets[wb1.SheetNames[0]];
      const df1 = XLSX.utils.sheet_to_json(ws1);

      // Map برای دسترسی سریع بر اساس کد عرضه
      const df1Map = {};
      df1.forEach((row) => {
        const code = String(row["کد عرضه"] || "").trim();
        if (code) df1Map[code] = row;
      });

      const order = [...new Set(df1.map((row) => String(row["کد عرضه"])))].filter(
        (x) => x && x !== "undefined"
      );

      // بررسی وجود ستون های مقدار عرضه/مقدار تقاضا در فایل اول
      const df1HasOfferQty = df1.some((r) => Object.prototype.hasOwnProperty.call(r, "مقدار عرضه"));
      const df1HasDemandQty = df1.some((r) => Object.prototype.hasOwnProperty.call(r, "مقدار تقاضا"));

      // --- فایل دوم (داده‌ها)
      const data2 = await file2.arrayBuffer();
      const wb2 = XLSX.read(data2, { type: "array" });
      const ws2 = wb2.Sheets[wb2.SheetNames[0]];
      const df2 = XLSX.utils.sheet_to_json(ws2);

      // --- ستون‌های خروجی پیشنهادی (به ترتیب از راست به چپ)
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

      // تصمیم‌گیری درباره اینکه چه ستون‌هایی در خروجی قرار بگیرند:
      // - برای 'عرضه' و 'تقاضا' اگر در فایل اول مقدار متناظر وجود داشته باشد، آن‌ها را بیاور.
      // - برای بقیه ستون‌ها اگر در فایل دوم موجود باشند، بیاور.
      const sampleRowDf2 = df2[0] || {};
      const availableCols = keepColumns.filter((col) => {
        if (col === "عرضه" && df1HasOfferQty) return true;
        if (col === "تقاضا" && df1HasDemandQty) return true;
        return Object.prototype.hasOwnProperty.call(sampleRowDf2, col);
      });

      // --- مرتب‌سازی بر اساس کد عرضه فایل اول و ساخت ردیف‌ها
      let result = [];
      order.forEach((code) => {
        const subset = df2
          .filter((row) => String(row["کد عرضه"]) === code)
          .map((row) => {
            let filtered = {};
            availableCols.forEach((col) => {
              let val = "";
              // اگر از فایل اول باید مقدار گرفته شود (مقدار عرضه / مقدار تقاضا)
              if (col === "عرضه" && df1HasOfferQty) {
                val = df1Map[code] ? df1Map[code]["مقدار عرضه"] : row[col] || "";
              } else if (col === "تقاضا" && df1HasDemandQty) {
                val = df1Map[code] ? df1Map[code]["مقدار تقاضا"] : row[col] || "";
              } else {
                val = row[col] || "";
              }

              // اگر مقداری شبیه عدد بود، آن را به Number تبدیل کن تا قالب عددی در اکسل اعمال شود
              if (isNumeric(val)) {
                // remove commas then convert
                const num = Number(String(val).replace(/,/g, ""));
                filtered[col] = num;
              } else {
                filtered[col] = val;
              }
            });
            return filtered;
          });

        if (subset.length > 0) {
          result.push(...subset);
          result.push({}); // ردیف خالی بین گروه‌ها
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
        sheet.addRow(availableCols.map((c) => (row && Object.prototype.hasOwnProperty.call(row, c) ? row[c] : "")));
      });

      // استایل‌دهی همه سلول‌ها (فونت، تراز، حاشیه، فرمت عدد)
      // توجه: ExcelJS فقط نام فونت را تنظیم می‌کند؛ برای اینکه فونت واقعاً نمایش داده شود باید روی سیستم مقصد نصب باشد.
      const fontName = "B Nazanin"; // اگر می‌خواهی به تاهما تغییر بدیم بگو
      sheet.eachRow((row) => {
        row.eachCell((cell) => {
          cell.font = { name: fontName, size: 12 };
          cell.alignment = { vertical: "middle", horizontal: "center", wrapText: true };
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
          // پس‌زمینه سفید (همونطور که قبلا بود)
          cell.fill = {
            type: "pattern",
            pattern: "solid",
            fgColor: { argb: "FFFFFFFF" },
          };
        });
      });

      // --- AutoFit بهتر برای فونت‌های فارسی
      // ضریب عرض برای فونت فارسی (در صورت نیاز مقدار را کمتر/بیشتر کن)
      const fontWidthFactor = 1.6; // عدد قابل تغییر: 1.2..1.8 بسته به فونت و اندازه
      sheet.columns.forEach((column) => {
        let maxLength = 0;
        column.eachCell({ includeEmpty: true }, (cell) => {
          // مقدار نمایشی سلول را بگیر
          let val = "";
          if (cell.value === null || cell.value === undefined) val = "";
          else if (typeof cell.value === "object" && cell.value.richText) {
            // اگر richText بود، جمع متن‌ها
            val = cell.value.richText.map((t) => t.text).join("");
          } else {
            val = String(cell.value);
          }
          // طول رشته
          const len = val.length;
          if (len > maxLength) maxLength = len;
        });

        // محاسبه عرض نهایی
        const calculated = Math.ceil(maxLength * fontWidthFactor) + 2; // padding
        column.width = Math.max(10, calculated); // حداقل عرض
      });

      // خروجی فایل
      const buffer = await workbook.xlsx.writeBuffer();
      saveAs(new Blob([buffer]), "اکسل_مرتب.xlsx");
      showAlert("فایل آماده شد و دانلود شد.");
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
