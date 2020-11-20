import xlsx from 'xlsx';

let ecpayExcelContent = [],
  mofExcelContent = [];

const handleECPayFileRead = (e) => {
  const files = e.target.files,
    f = files[0];
  const reader = new FileReader();

  reader.onload = function (e) {
    try {
      const excelFile = new Uint8Array(e.target.result);
      const workbook = xlsx.read(excelFile, { type: 'array' });
      ecpayExcelContent = xlsx.utils.sheet_to_json(
        workbook.Sheets[workbook.SheetNames[0]],
      );
    } catch (err) {
      console.error(err);
      alert('讀取失敗，請檢查檔案格式是否正確');
      ecpayExcelContent = [];
    }
  };
  reader.readAsArrayBuffer(f);
};

const handleMOFFileRead = (e) => {
  const files = e.target.files,
    f = files[0];
  const reader = new FileReader();

  reader.onload = function (e) {
    try {
      const excelFile = new Uint8Array(e.target.result);
      const workbook = xlsx.read(excelFile, { type: 'array' });
      mofExcelContent = xlsx.utils.sheet_to_json(
        workbook.Sheets[workbook.SheetNames[0]],
      );
    } catch (err) {
      console.error(err);
      alert('讀取失敗，請檢查檔案格式是否正確');
      mofExcelContent = [];
    }
  };
  reader.readAsArrayBuffer(f);
};

const handleMergeAndDownload = () => {
  const mergedExcel = mergeExcelContent(mofExcelContent, ecpayExcelContent);
  downloadExcelContent(mergedExcel);
};

const mergeExcelContent = (mofExcelContent, ecpayExcelContent) => {
  const mergedExcelContent = mofExcelContent.map((mofInvoice) => {
    const ecpayInvoice = ecpayExcelContent.find(
      (ei) => ei['發票號碼'] === mofInvoice['發票號碼'],
    );
    if (!ecpayInvoice) return mofInvoice;

    return {
      ...mofInvoice,
      客戶姓名: ecpayInvoice['客戶姓名'],
      客戶地址: ecpayInvoice['客戶地址'],
      客戶手機號碼: ecpayInvoice['客戶手機號碼'],
    };
  });

  return mergedExcelContent;
};

const downloadExcelContent = (mergedExcelContent) => {
  const sheetHeader = [
    '發票號碼',
    '註記欄(不轉入進銷項媒體申報檔)',
    '格式代號',
    '發票狀態',
    '發票日期',
    '買方統一編號',
    '買方名稱',
    '賣方統一編號',
    '賣方名稱',
    '寄送日期',
    '銷售額合計',
    '零稅銷售額',
    '免稅銷售額',
    '營業稅',
    '總計',
    '課稅別',
    '載具類別編號',
    '載具編碼',
    '客戶姓名',
    '客戶手機號碼',
    '客戶地址',
  ];
  const sheet = xlsx.utils.json_to_sheet(mergedExcelContent, {
    header: sheetHeader,
  });
  const workbook = xlsx.utils.book_new();
  xlsx.utils.book_append_sheet(workbook, sheet, 'Sheet0');
  xlsx.writeFile(workbook, 'merged.xlsx', { bookType: 'xlsx' });
};

document
  .querySelector('#ECPayInvoiceCSV')
  .addEventListener('change', handleECPayFileRead, false);

document
  .querySelector('#MOFInvoiceExcel')
  .addEventListener('change', handleMOFFileRead, false);

document
  .querySelector('#startMerge')
  .addEventListener('click', handleMergeAndDownload, false);
