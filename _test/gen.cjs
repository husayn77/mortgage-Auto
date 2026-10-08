// Builds a sample input workbook (sample-input.xlsx) to test the mapper via the
// running Next.js server.
const ExcelJS = require("exceljs");
const path = require("path");

async function main() {
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet("Data");
  ws.addRow([
    "SN",
    "Account Number",
    "Account Name",
    "PIN Number",
    "Balance",
    "Address",
    "Next of Kin",
    "NIN",
    "BVN",
    "Phone",
    "DOB",
    "Employer",
    "Sgns",
  ]);
  ws.addRow([
    "ACCESS",
    1000501578,
    "ISA MUHAMMED",
    "PEN100907682416",
    4155533.8704,
    "JUNIO STAFF QUARTERS ABAJI",
    "YUSUF MUHAMMED YUSUF",
    7544345667,
    11122222111,
    8036097918,
    "01/09/1987",
    "ACCESS BANK",
    "Signatory A",
  ]);
  ws.addRow([
    "ACCESS",
    1000501579,
    "UMMUHANI ISAH OZIOHU",
    "PEN200779658725",
    3294795.3885,
    "ALEYTA OPPOSITE ECWA CHURCH, LUGBE, ABUJA",
    "MARYAM ISAH",
    2345678999,
    234567899,
    8060539893,
    "15/01/1975",
    "ACCESS BANK",
    "Signatory B",
  ]);
  const out = path.join(__dirname, "sample-input.xlsx");
  await wb.xlsx.writeFile(out);
  console.log("Wrote", out);
}

main().catch((e) => {
  console.error(e);
  process.exit(1);
});
