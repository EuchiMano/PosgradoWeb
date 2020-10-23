using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Web;

namespace PosgradoWeb.Controller
{
    public class clsExcel : IDisposable
    {
        public List<List<string>> mtdConvertirExcel(byte[] luFile, string puName = null)
        {
            System.IO.MemoryStream fileXls = new MemoryStream(luFile);
            List<List<string>> luResult = new List<List<string>>();

            try
            {
                using (SpreadsheetDocument doc = SpreadsheetDocument.Open(fileXls, false))
                {
                    //Read the first Sheet from Excel file.
                    Sheet sheet = new Sheet();
                    if (puName == null) sheet = doc.WorkbookPart.Workbook.Sheets.GetFirstChild<Sheet>();
                    else
                    {
                        Sheets luHojas = doc.WorkbookPart.Workbook.Sheets;
                        foreach (Sheet hoja in luHojas)
                        {
                            if (hoja.Name.ToString().ToUpper() == puName.ToUpper())
                            {
                                sheet = hoja;
                            }
                        }
                    }

                    SharedStringTablePart sstpart = doc.WorkbookPart.GetPartsOfType<SharedStringTablePart>().First();
                    SharedStringTable sst = sstpart.SharedStringTable;

                    //Get the Worksheet instance.
                    Worksheet worksheet = (doc.WorkbookPart.GetPartById(sheet.Id.Value) as WorksheetPart).Worksheet;

                    //Fetch all the rows present in the Worksheet.
                    IEnumerable<Row> rows = worksheet.GetFirstChild<SheetData>().Descendants<Row>();

                    //Create a new DataTable.
                    luResult = new List<List<string>>();

                    //Loop through the Worksheet rows.

                    foreach (Row row in rows)
                    {
                        //Use the first row to add columns to DataTable.
                        List<string> luFila = new List<string>();
                        int columnIndex = 0;
                        int cont = 0;
                        foreach (Cell cell in row.Descendants<Cell>())
                        {
                            for (int c = 0; c <= cont; c++)
                            {
                                int cellColumnIndex = (int)GetColumnIndexFromName(GetColumnName(cell.CellReference));
                                cellColumnIndex--;
                                if (columnIndex < cellColumnIndex)
                                {
                                    do
                                    {

                                        //luFila[columnIndex] = null;
                                        luFila.Add(null);
                                        columnIndex++;
                                    }
                                    while (columnIndex < cellColumnIndex);

                                }
                                if (cell.DataType != null && cell.DataType.Value == CellValues.SharedString)
                                {
                                    string lsValor = sst.ChildElements[Convert.ToInt32(cell.CellValue.Text)].InnerText;
                                    luFila.Add(lsValor);

                                    //luFila[columnIndex] = lsValor;
                                }
                                else
                                     if (cell.CellValue != null)
                                {
                                    luFila.Add(cell.CellValue.Text);
                                }
                                else if (cell.CellValue == null)
                                {
                                    luFila.Add(null);
                                }
                                columnIndex++;
                            }
                        }
                        if (luFila.Count > 0)
                        {
                            luResult.Add(luFila);
                            cont = luResult[0].Count;
                        }

                        //#region "MI codigo anterior"
                        //foreach (Cell cell in row.Descendants<Cell>())
                        //{
                        //    if (cell.DataType != null && cell.DataType.Value == CellValues.SharedString)
                        //    {

                        //        //string lsValor = sst.ChildElements(Convert.ToInt32(cell.CellValue.Text)).InnerText;
                        //        string lsValor = sst.ChildElements[Convert.ToInt32(cell.CellValue.Text)].InnerText;

                        //        luFila.Add(lsValor);
                        //    }
                        //    else
                        //    if (cell.CellValue != null)
                        //    {
                        //        luFila.Add(cell.CellValue.Text);
                        //    }
                        //    else if (cell.CellValue == null)
                        //    {
                        //        luFila.Add(null);
                        //    }


                        //}

                        //if (luFila.Count > 0)
                        //{
                        //    luResult.Add(luFila);
                        //}
                        //#endregion

                    }

                }

            }
            catch (Exception ex)
            {
                return null;
            }
            return luResult;

        }



        private static int? GetColumnIndexFromName(string columnName)
        {
            //return columnIndex;
            string name = columnName;
            int number = 0;
            int pow = 1;
            for (int i = name.Length - 1; i >= 0; i--)
            {
                number += (name[i] - 'A' + 1) * pow;
                pow *= 26;
            }
            return number;
        }

        private static string GetColumnName(string cellReference)
        {
            // Create a regular expression to match the column name portion of the cell name.
            Regex regex = new Regex("[A-Za-z]+");
            Match match = regex.Match(cellReference);

            return match.Value;
        }


        #region IDisposable Support
        // some fields that require cleanup
        private bool disposed = false; // to detect redundant calls

        protected virtual void Dispose(bool disposing)
        {
            if (!disposed)
            {
                // Dispose unmanaged managed resources.
                disposed = true;
            }
        }
        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }
        #endregion
    }
}