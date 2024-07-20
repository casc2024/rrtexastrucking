using Microsoft.Ajax.Utilities;
using Newtonsoft.Json;
using NPOI.HSSF.UserModel;
using NPOI.POIFS.Crypt.Dsig;
using NPOI.SS.Formula.Functions;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Web;
using System.Web.Mvc;
using updateTemplateWebv2.Models;

namespace updateTemplateWebv2.Controllers
{
    public class HomeController : Controller
    {
        public ActionResult Index()
        {
            return View();
        }

        [HttpPost]
        public ActionResult UploadExcel(HttpPostedFileBase file)
        {
            if (file != null && file.ContentLength > 0 && (file.FileName.EndsWith(".xlsx") || file.FileName.EndsWith(".xls")))
            {
                var data = ReadExcelFile(file.InputStream);
                var jsonData = JsonConvert.SerializeObject(data, Formatting.Indented);
                var listaInicial = JsonConvert.DeserializeObject<List<ModelData>>(jsonData);

                var pathTemplate = Path.Combine(Server.MapPath("~/File"), "InvoiceTemplate.xlsm");
                IWorkbook workbook;
                using (FileStream fileTemplate = new FileStream(pathTemplate, FileMode.Open, FileAccess.Read))
                {
                    workbook = new XSSFWorkbook(fileTemplate);
                }

                ISheet sheet1 = workbook.GetSheetAt(0);
                ISheet sheet2 = workbook.GetSheetAt(1);
                EditSheet(sheet2, listaInicial);

                List<string> listaOriginal = new List<string>();
                listaInicial.ForEach(x => {
                    if (!listaOriginal.Contains(x.ORIGEN)) { 
                        listaOriginal.Add(x.ORIGEN);
                    }
                });

                XSSFSheet _sheet1 = (XSSFSheet)workbook.GetSheet("Sheet1");
                CreateListOrigen(_sheet1, sheet2, listaOriginal);
                CreateTable(_sheet1, sheet2, listaInicial, listaOriginal, workbook);

                using (FileStream fileTemplate = new FileStream(pathTemplate, FileMode.Create, FileAccess.Write))
                {
                    workbook.Write(fileTemplate);
                }

            }
            return View();
        }

        private List<Dictionary<string, object>> ReadExcelFile(Stream fileStream)
        {
            IWorkbook workbook = new XSSFWorkbook(fileStream);
            ISheet sheet = workbook.GetSheetAt(0);

            var rows = new List<Dictionary<string, object>>();

            IRow headerRow = sheet.GetRow(0);
            int cellCount = headerRow.LastCellNum;

            for (int i = 1; i <= sheet.LastRowNum; i++)
            {
                IRow row = sheet.GetRow(i);
                var rowData = new Dictionary<string, object>();

                for (int j = 0; j < cellCount; j++)
                {
                    ICell cell = row.GetCell(j);
                    string header = headerRow.GetCell(j).ToString();
                    rowData[header] = GetCellValue(cell);
                }

                rows.Add(rowData);
            }

            return rows;
        }

        private object GetCellValue(ICell cell)
        {
            if (cell == null)
                return null;

            switch (cell.CellType)
            {
                case CellType.Boolean:
                    return cell.BooleanCellValue;
                case CellType.Numeric:
                    if (DateUtil.IsCellDateFormatted(cell))
                        return cell.DateCellValue;
                    else
                        return cell.NumericCellValue;
                case CellType.String:
                    return cell.StringCellValue;
                case CellType.Blank:
                    return null;
                default:
                    return cell.ToString();
            }
        }

        static void EditSheet(ISheet sheet, List<ModelData> listModelData)
        {
            int row = 0;
            foreach (var item in listModelData)
            {
                IRow newRow = sheet.CreateRow(row);
                newRow.CreateCell(0).SetCellValue($"{item.ORIGEN.TrimEnd()}{item.DESTINATION.TrimEnd()}{item.PRODUCT.TrimEnd()}");
                newRow.CreateCell(1).SetCellValue(item.ORIGEN.TrimEnd());
                newRow.CreateCell(2).SetCellValue(item.DESTINATION.TrimEnd());
                newRow.CreateCell(3).SetCellValue(item.PRODUCT.TrimEnd());
                newRow.CreateCell(4).SetCellValue(item.RATE);
                row++;
            }
            
            
        }

        static void CreateListOrigen(XSSFSheet sheet1, ISheet sheet2, List<string> listOrigen) {
            int row = 0;
            foreach (var item in listOrigen)
            {
                IRow rowSheet = sheet2.GetRow(row);
                rowSheet.CreateCell(5).SetCellValue(item);
                row++; 
            }           
            IDataValidationHelper validationHelper = new XSSFDataValidationHelper(sheet1);
            CellRangeAddressList addressList = new CellRangeAddressList(13, 199, 2, 2);
            IDataValidationConstraint constraint = validationHelper.CreateFormulaListConstraint($"=Sheet2!$F$1:$F${listOrigen.Count()}");
            IDataValidation dataValidation = validationHelper.CreateValidation(constraint, addressList);
            dataValidation.SuppressDropDownArrow = true;
            sheet1.AddValidationData(dataValidation);
        }

        static void CreateTable(XSSFSheet sheet1, ISheet sheet, List<ModelData> listModelData, List<string> listOrigen, IWorkbook workbook) { 
            Dictionary<string, List<string>> listaDestino = new Dictionary<string, List<string>>();
            string nombreLista = string.Empty;
            listOrigen.ForEach(x => {
                nombreLista = Regex.Replace(x, @"[^a-zA-Z0-9.]", "");
                var obj = listModelData.Where(z => z.ORIGEN == x).Select(z => z.DESTINATION).Distinct().ToList();
                listaDestino.Add($"{x}|{nombreLista}", obj);
            });

            int row = 0;
            int col = 6;
            listaDestino.ForEach(x => {
                IRow rowSheet = sheet.GetRow(row);
                rowSheet.CreateCell(col).SetCellValue(x.Key.Split('|')[0]);
                row++;
                rowSheet = sheet.GetRow(row);
                rowSheet.CreateCell(col).SetCellValue($"Table{x.Key.Split('|')[1]}");
                row++;
                x.Value.ForEach(z => {
                    rowSheet = sheet.GetRow(row);
                    rowSheet.CreateCell(col).SetCellValue(z);
                    row++;
                });

                col++;
                IName namedRange = workbook.CreateName();
                namedRange.NameName = $"Table{x.Key.Split('|')[1]}";
                var letterCol = GetColumnLetter(col - 1);
                namedRange.RefersToFormula = $"Sheet2!${letterCol}$1:${letterCol}${row + 1}";
                row = 0;
            });
            
            IDataValidationHelper validationHelper = new XSSFDataValidationHelper(sheet1);
            CellRangeAddressList addressList = new CellRangeAddressList(13, 13, 3, 3);
            IDataValidationConstraint constraint = validationHelper.CreateFormulaListConstraint($"=INDIRECTO(BUSCARH(C14,Sheet2!$G$1:$AB$16,2))");
            IDataValidation dataValidation = validationHelper.CreateValidation(constraint, addressList);
            dataValidation.SuppressDropDownArrow = true;
            sheet1.AddValidationData(dataValidation);
        }

        static string GetColumnLetter(int columnIndex)
        {
            string columnLetter = string.Empty;
            while (columnIndex >= 0)
            {
                columnLetter = (char)('A' + (columnIndex % 26)) + columnLetter;
                columnIndex = columnIndex / 26 - 1;
            }
            return columnLetter;
        }
    }
}