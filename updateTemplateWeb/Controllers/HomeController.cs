using Newtonsoft.Json;
using OfficeOpenXml;
using OfficeOpenXml.DataValidation;
using Serilog;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Web;
using System.Web.Mvc;
using System.Windows;
using updateTemplateWeb.Models;

namespace updateTemplateWeb.Controllers
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
            string fileNameDownload = string.Empty;
            string contentType = string.Empty;
            byte[] fileBytes = null;

            try
            {
                if (file != null && file.ContentLength > 0 && (file.FileName.EndsWith(".xlsx") || file.FileName.EndsWith(".xls")))
                {
                    var fileName = Path.GetFileName(file.FileName);
                    var path = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "App_Data/Uploads", fileName);
                    file.SaveAs(path);

                    DataTable dataTable = ReadExcelFile(path);
                    string json = ConvertDataTableToJson(dataTable);
                    var listaInicial = JsonConvert.DeserializeObject<List<ModelData>>(json);

                    var pathTemplate = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "File", "InvoiceTemplate.xlsm");

                    using (var package = new ExcelPackage(new FileInfo(pathTemplate)))
                    {
                        ExcelWorksheet worksheet = package.Workbook.Worksheets[0];

                        //limpiar validaciones de la primera hoja
                        ClearValidatation("C", 14, 200, worksheet);
                        ClearValidatation("D", 14, 200, worksheet);
                        ClearValidatation("E", 14, 200, worksheet);

                        ExcelWorksheet newWorksheet = package.Workbook.Worksheets[1];
                        if (newWorksheet != null)
                        {
                            package.Workbook.Worksheets.Delete(newWorksheet);
                            package.Workbook.Worksheets.Add("Sheet2");
                        }
                        newWorksheet = package.Workbook.Worksheets[1];
                        for (int i = 0; i < listaInicial.Count(); i++)
                        {
                            newWorksheet.Cells[i + 1, 1].Value = $"{listaInicial[i].ORIGEN.TrimEnd()}{listaInicial[i].DESTINATION.TrimEnd()}{listaInicial[i].PRODUCT.TrimEnd()}";
                            newWorksheet.Cells[i + 1, 2].Value = listaInicial[i].ORIGEN.TrimEnd();
                            newWorksheet.Cells[i + 1, 3].Value = listaInicial[i].DESTINATION.TrimEnd();
                            newWorksheet.Cells[i + 1, 4].Value = listaInicial[i].PRODUCT.TrimEnd();
                            newWorksheet.Cells[i + 1, 5].Value = listaInicial[i].RATE;
                        }

                        List<string> listaOrigen = new List<string>();
                        listaInicial.ForEach(x =>
                        {
                            listaOrigen.Add(x.ORIGEN.TrimEnd());
                        });

                        CreateDropDownList(6, "F", "C", worksheet, "Sheet2", newWorksheet, listaOrigen.Distinct().ToList(), 14, 200);
                        //CreateNestedDropDownList("C","D", worksheet, "Sheet2", newWorksheet, listaInicial, 14, 200);
                        //Agregar columas del destino
                        List<string> listaDestino = new List<string>();
                        listaInicial.ForEach(x =>
                        {
                            listaDestino.Add(x.DESTINATION.TrimEnd());
                        });
                        CreateDropDownList(7, "G", "D", worksheet, "Sheet2", newWorksheet, listaDestino.Distinct().ToList(), 14, 200);

                        //Agregar columas del material
                        List<string> listaMaterial = new List<string>();
                        listaInicial.ForEach(x =>
                        {
                            listaMaterial.Add(x.PRODUCT.TrimEnd());
                        });
                        CreateDropDownList(8, "H", "E", worksheet, "Sheet2", newWorksheet, listaMaterial.Distinct().ToList(), 14, 200);

                        for (int i = 14; i <= 200; i++)
                        {
                            worksheet.Cells[$"F{i}"].Formula = $"IFNA(VLOOKUP(C{i}&D{i}&E{i},Sheet2!$A$1:$E${listaInicial.Count()},5,FALSE),0)";
                        }

                        newWorksheet.Hidden = eWorkSheetHidden.Hidden;

                        worksheet.Protection.SetPassword("150683");
                        worksheet.Protection.IsProtected = true;
                        worksheet.Protection.AllowSelectLockedCells = false;

                        newWorksheet.Protection.SetPassword("150683");
                        newWorksheet.Protection.IsProtected = true;
                        newWorksheet.Protection.AllowSelectLockedCells = false;

                        package.Workbook.Protection.LockStructure = true;
                        package.Workbook.Protection.SetPassword("150683");

                        package.Save();
                    }

                    fileNameDownload = Path.GetFileName(pathTemplate);
                    contentType = MimeMapping.GetMimeMapping(pathTemplate);
                    fileBytes = System.IO.File.ReadAllBytes(pathTemplate);
                }
                ViewBag.MessageType = "1";                

                //return File(fileBytes, contentType, fileNameDownload);
                //RedirectToAction("DownloadFile", new { fileName = "InvoiceTemplate.xlsm" });
                
            }
            catch (Exception ex)
            {
                ViewBag.MessageType = "0";
                ViewBag.Message = $"An error occurred in the operation";
                Log.Error($"Ocurrió un error: {ex.Message}");                
            }   
            return View("Index");
        }

        [HttpPost]
        public JsonResult GenerateFile()
        {
            var fileName = "InvoiceTemplate.xlsm";
            return Json(new { fileUrl = Url.Action("DownloadFile", "Home", new { fileName }) });
        }

        public ActionResult DownloadFile(string fileName)
        {
            var filePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "File", fileName);
            byte[] fileBytes = System.IO.File.ReadAllBytes(filePath);
            string fileType = "application/octet-stream";

            return File(fileBytes, fileType, fileName);
        }

        //public ActionResult DownloadFile(string fileName)
        //{
        //    //var fileName = "InvoiceTemplate.xlsm";
        //    var pathTemplate = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "File", "InvoiceTemplate.xlsm");

        //    byte[] fileBytes = System.IO.File.ReadAllBytes(pathTemplate);
        //    return File(fileBytes, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", fileName);
        //}

        private static void ClearValidatation(string nombreColumna, int filaInicio, int filaFin, ExcelWorksheet worksheet) {
            for (int i = filaInicio; i <= filaFin; i++)
            {                
                var validation = worksheet.DataValidations[$"{nombreColumna}{i}"];
                worksheet.DataValidations.Remove(validation);
            }
        }

        private static string ConvertDataTableToJson(DataTable dataTable)
        {
            return JsonConvert.SerializeObject(dataTable, Formatting.Indented);
        }

        private static DataTable ReadExcelFile(string filePath)
        {

            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            using (ExcelPackage package = new ExcelPackage(new FileInfo(filePath)))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets[0];
                DataTable dataTable = new DataTable();

                for (int col = 1; col <= worksheet.Dimension.End.Column; col++)
                {
                    dataTable.Columns.Add(worksheet.Cells[1, col].Text);
                }

                for (int row = 2; row <= worksheet.Dimension.End.Row; row++)
                {
                    DataRow dataRow = dataTable.NewRow();
                    for (int col = 1; col <= worksheet.Dimension.End.Column; col++)
                    {
                        dataRow[col - 1] = worksheet.Cells[row, col].Text;
                    }
                    dataTable.Rows.Add(dataRow);
                }

                return dataTable;
            }
        }

        private static void CreateDropDownList(int col, string nomCol, string colListaDesplegable, ExcelWorksheet worksheet, string nombreHoja, ExcelWorksheet newWorksheet, List<string> lista, int filaInicio, int filaFin)
        {

            string[] dropdownValues = lista.ToArray();

            for (int i = 0; i < dropdownValues.Length; i++)
            {
                newWorksheet.Cells[i + 1, col].Value = dropdownValues[i];
            }
            string address = $"{nombreHoja}!{nomCol}1:{nomCol}{dropdownValues.Length}";

            for (int i = filaInicio; i <= filaFin; i++)
            {
                var validationRange = worksheet.Cells[$"{colListaDesplegable}{i}"];
                var validation = worksheet.DataValidations.AddListValidation(validationRange.Address);
                validation.Formula.ExcelFormula = address;
            }
        }

        private static void CreateNestedDropDownList(string colListDesplegableDependiente,string colListaDesplegable, ExcelWorksheet worksheet, string nombreHoja, ExcelWorksheet newWorksheet, List<ModelData> listaOriginal, int filaInicio, int filaFin)
        {

            string[] dropdownValues;
            List<string> listaPendiente = new List<string>();
            List<string> listaIndependiente = new List<string>();
            List<string> listaData = new List<string>();

            listaOriginal.ForEach(x =>
            {
                if (!listaIndependiente.Contains(x.ORIGEN)) {
                    listaIndependiente.Add(x.ORIGEN);
                }                
            });
            int colInicial = 7;
            string nombreLista = string.Empty;
            var destinatarios = new Dictionary<string, List<string>>();
            foreach (var item in listaIndependiente)
            {
                listaData = listaOriginal.Where(z => z.ORIGEN == item).Select(z => z.DESTINATION).Distinct().ToList();
                listaPendiente = new List<string>();
                listaPendiente.AddRange(listaData);
                nombreLista = Regex.Replace(item, @"[^a-zA-Z0-9.]", "");
                destinatarios.Add($"{nombreLista}", listaPendiente);
            }

            int row = 1;            
            foreach (var destinatario in destinatarios)
            {                
                newWorksheet.Cells[row, colInicial].Value = destinatario.Key;               
                destinatario.Value.ForEach(x =>
                {
                    newWorksheet.Cells[row + 1, colInicial].Value = x;
                    row++;
                });

                var range = newWorksheet.Cells[1, colInicial, destinatarios[destinatario.Key].Count + 1, colInicial];
                var table = newWorksheet.Tables.Add(range, destinatario.Key);
                table.ShowHeader = true;
                colInicial++;
                row = 1;                               
            }

            string address = "=INDIRECTO(BUSCARH(C14,Sheet2!$G$1:$AB$16,2))";
            var validationRange = worksheet.Cells[$"D14"];
            var itemValidation = worksheet.DataValidations.AddListValidation(validationRange.Address);
            itemValidation.Formula.ExcelFormula = address;
            itemValidation.AllowBlank = true;
            itemValidation.ShowErrorMessage = true;
            itemValidation.ErrorStyle = ExcelDataValidationWarningStyle.stop;
            itemValidation.ErrorTitle = "Error de validación";
            itemValidation.Error = "Selecciona un valor válido";

            //for (int i = filaInicio; i <= filaFin; i++) {
            //    var validationRange = worksheet.Cells[$"{colListaDesplegable}{i}"];
            //    var itemValidation = worksheet.DataValidations.AddListValidation(validationRange.Address);
            //    var cellValueC14 = worksheet.Cells[$"{colListDesplegableDependiente}{i}"].Text;                
            //    itemValidation.Formula.ExcelFormula =string.Format("INDIRECTO(BUSCARH({0}{1},Sheet2!$G$1:$AB$16,2))", colListDesplegableDependiente, i);
            //    itemValidation.AllowBlank = true;
            //}

            //for (int i = 0; i < dropdownValues.Length; i++)
            //{
            //    newWorksheet.Cells[i + 1, col].Value = dropdownValues[i];
            //}

            //string address = $"{nombreHoja}!{nomCol}1:{nomCol}{dropdownValues.Length}";

            //for (int i = filaInicio; i <= filaFin; i++)
            //{
            //    var validationRange = worksheet.Cells[$"{colListaDesplegable}{i}"];
            //    var validation = worksheet.DataValidations.AddListValidation(validationRange.Address);
            //    validation.Formula.ExcelFormula = address;
            //}
        }

        public static string GetExcelColumnName(int columnNumber)
        {
            int dividend = columnNumber;
            string columnName = String.Empty;
            int modulo;

            while (dividend > 0)
            {
                modulo = (dividend - 1) % 26;
                columnName = Convert.ToChar(65 + modulo) + columnName;
                dividend = (dividend - modulo) / 26;
            }

            return columnName;
        }
    }
}