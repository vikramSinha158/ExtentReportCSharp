using System;
using System.Collections;
using System.IO;
using System.Runtime.InteropServices;
using xl = Microsoft.Office.Interop.Excel;

namespace R1.Automation.DataReadWrite.Core.Excel
{
    public class ExcelReadingWriting
    {
        xl.Application xlApp = null;
        xl.Workbooks workbooks = null;
        xl.Workbook workbook = null;
        Hashtable sheets;
        int colNumber = 0;
        int rowNumber = 0;
        xl.Worksheet worksheet = null;
        xl.Range range = null;



        /// <summary>This method is used for open excel file</summary>
        /// <param name="xlFilePath"></param>
        public void OpenExcel(string xlFilePath)
        {
            xlApp = new xl.Application();
            workbooks = xlApp.Workbooks;
            try
            {
                workbook = workbooks.Open(xlFilePath);
            }catch(Exception e)
            {
                throw new DirectoryNotFoundException("xlFile Path is not Found");
            }
            sheets = new Hashtable();
            int count = 1;
            // Storing worksheet names in Hashtable.
            foreach (xl.Worksheet sheet in workbook.Sheets)
            {
                sheets[count] = sheet.Name;
                count++;
            }
        }

        /// <summary>This method is used for close excel file</summary>
        /// <param name="xlFilePath"></param>
        public void CloseExcel(string xlFilePath)
        {
            workbook.Close(false, xlFilePath, null); // Close the connection to workbook
            Marshal.FinalReleaseComObject(workbook); // Release unmanaged object references.
            workbook = null;

            workbooks.Close();
            Marshal.FinalReleaseComObject(workbooks);
            workbooks = null;

            xlApp.Quit();
            Marshal.FinalReleaseComObject(xlApp);
            xlApp = null;
        }

        /// <summary>This method is used for reading data from excel file</summary>
        /// <param name="xlFilePath"></param>
        /// <param name="sheetName"></param>
        /// <param name="colName"></param>
        /// <param name="rowName"></param>
        /// <returns>It returns string value</returns>
        public string GetCellData(string xlFilePath, string sheetName, string colName, string rowName)
        {
            OpenExcel(xlFilePath);

            string value = string.Empty;
            colNumber = 0;
            rowNumber = 0;

            try
            {
                FindRowNumAndColumnNum(sheetName, colName, rowName);

                if (colNumber > 0 && rowNumber > 0)
                {
                    value = Convert.ToString((range.Cells[rowNumber, colNumber] as xl.Range).Value2);
                }
                Marshal.FinalReleaseComObject(worksheet);
                worksheet = null;
            }
            finally
            {
                CloseExcel(xlFilePath);
            }
            return value;
        }

        /// <summary>This method is used for writing data in excel file</summary>
        /// <param name="xlFilePath"></param>
        /// <param name="sheetName"></param>
        /// <param name="colName"></param>
        /// <param name="rowName"></param>
        /// <param name="value"></param>
        /// <returns>It returns boolean value</returns>
        public bool SetCellData(string xlFilePath, string sheetName, string colName, string rowName, string value)
        {
            OpenExcel(xlFilePath);
            colNumber = 0;
            rowNumber = 0;


            try
            {
                    FindRowNumAndColumnNum(sheetName, colName, rowName);

                    if (colNumber > 0 && rowNumber > 0)
                    {
                        range.Cells[rowNumber, colNumber] = value;
                        workbook.Save();
                    }
                    Marshal.FinalReleaseComObject(worksheet);
                    worksheet = null;

             }
            catch (Exception ex)
            {
                return false;
            }
            finally
            {
                CloseExcel(xlFilePath);
            }
            return true;
        }

        /// <summary>This method is used for Find Row Num And Column Num from Excel</summary>
        /// <param name="sheetName"></param>
        /// <param name="colName"></param>
        /// <param name="rowName"></param>
        private void FindRowNumAndColumnNum(string sheetName, string colName, string rowName)
        {

            int sheetValue = 0;



                if (sheets.ContainsValue(sheetName))
                {
                    foreach (DictionaryEntry sheet in sheets)
                    {
                        if (sheet.Value.Equals(sheetName))
                        {
                            sheetValue = (int)sheet.Key;
                        }
                    }

                    
                    worksheet = workbook.Worksheets[sheetValue] as xl.Worksheet;
                    range = worksheet.UsedRange;

                    for (int i = 1; i <= range.Columns.Count; i++)
                    {
                        string colNameValue = Convert.ToString((range.Cells[1, i] as xl.Range).Value2);
                        if (colNameValue.ToLower() == colName.ToLower())
                        {
                            colNumber = i;
                            break;
                        }
                    }

                    for (int i = 1; i <= range.Rows.Count; i++)
                    {
                        string rawNameValue = Convert.ToString((range.Cells[i, 1] as xl.Range).Value2);

                        if (rawNameValue.ToLower() == rowName.ToLower())
                        {
                            rowNumber = i;
                            break;
                        }
                    }

                }
        }
    }
}