using ExcelDataReader;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Web;

namespace OutPayslip.Services
{
    public class ExcelReader
    {
        public static DataSet CreateDataset(String inputFilePath)
        {
            DataSet finalDataSet = new DataSet();
            DataSet excelDataSet = ExcelToDataSet(inputFilePath);
            if (excelDataSet != null && excelDataSet.Tables.Count > 0)
            {
                foreach (DataTable excelDataTable in excelDataSet.Tables)
                {
                    if (excelDataTable.Rows != null && excelDataTable.Rows.Count > 0)
                    {
                        var rows = excelDataTable.Rows.Cast<DataRow>().Where(row => !row.ItemArray.All(field => field is DBNull));
                        if (rows != null && rows.Any())
                        {
                            DataTable currentDataTable = rows.CopyToDataTable();
                            currentDataTable.TableName = excelDataTable.TableName;
                            finalDataSet.Tables.Add(currentDataTable);
                        }
                    }
                }
            }
            return finalDataSet;
        }
        public static DataSet ExcelToDataSet(string pathToExcel)
        {

            DataSet excelPages = null;

            IExcelDataReader excelReader = default(IExcelDataReader);

            using (FileStream excelStream = new FileStream(pathToExcel, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                if (Path.GetExtension(pathToExcel).ToLower() == ".xls")
                {
                    excelReader = ExcelReaderFactory.CreateBinaryReader(excelStream);
                }
                else
                {
                    excelReader = ExcelReaderFactory.CreateOpenXmlReader(excelStream);
                }
                excelPages = excelReader.AsDataSet(new ExcelDataSetConfiguration()
                {
                    UseColumnDataType = true,
                    ConfigureDataTable = (tableReader) => new ExcelDataTableConfiguration()
                    {
                        UseHeaderRow = true,
                    }
                });
            }
            return excelPages;
        }

    }
}