using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Microsoft.WindowsAzure.Storage.Auth;
using Microsoft.WindowsAzure.Storage.Blob;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Web;

namespace CSWebAppAzureExcelImportExport
{
    public class Helper
    {
        private static readonly StorageCredentials cred = new StorageCredentials("[Your storage account name]", "[Your storage account key]");
        private static readonly CloudBlobContainer container = new CloudBlobContainer(new Uri("http://[Your storage account name].blob.core.windows.net/[Your container name] /"), cred);
        private static readonly string connectionStr = "Azure SQL Server Connection String";

        private static readonly string directoryPath = $"{ AppDomain.CurrentDomain.BaseDirectory}\\Downloads";
        private static readonly string excelName = "Student.xlsx";
        private static readonly List<string> columns = new List<string>() { "Name", "Class", "Score", "Sex" };
        private static readonly string tableName = "StudentScore";

        public static string DBExportToExcel()
        {
            string result = string.Empty;
            try
            {
                //Get datatable from db
                DataSet ds = new DataSet();
                SqlConnection connection = new SqlConnection(connectionStr);                
                SqlCommand cmd = new SqlCommand($"SELECT {string.Join(",", columns)} FROM {tableName}", connection);
                using (SqlDataAdapter adapter = new SqlDataAdapter(cmd))
                {
                    adapter.Fill(ds);
                }
                //Check directory
                if (!Directory.Exists(directoryPath))
                {
                    Directory.CreateDirectory(directoryPath);
                }
                // Delete the file if it exists
                string filePath = $"{directoryPath}//{excelName}";
                if (File.Exists(filePath))
                {
                    File.Delete(filePath);
                }

                if (ds.Tables.Count > 0 && ds.Tables[0] != null || ds.Tables[0].Columns.Count > 0)
                {
                    DataTable table = ds.Tables[0];

                    using (var spreadsheetDocument = SpreadsheetDocument.Create(filePath, SpreadsheetDocumentType.Workbook))
                    {
                        // Create SpreadsheetDocument
                        WorkbookPart workbookPart = spreadsheetDocument.AddWorkbookPart();
                        workbookPart.Workbook = new Workbook();
                        var sheetPart = spreadsheetDocument.WorkbookPart.AddNewPart<WorksheetPart>();
                        var sheetData = new SheetData();
                        sheetPart.Worksheet = new Worksheet(sheetData);
                        Sheets sheets = spreadsheetDocument.WorkbookPart.Workbook.AppendChild<Sheets>(new Sheets());
                        string relationshipId = spreadsheetDocument.WorkbookPart.GetIdOfPart(sheetPart);
                        Sheet sheet = new Sheet() { Id = relationshipId, SheetId = 1, Name = table.TableName };
                        sheets.Append(sheet);

                        //Add header to sheetData
                        Row headerRow = new Row();
                        List<String> columns = new List<string>();
                        foreach (DataColumn column in table.Columns)
                        {
                            columns.Add(column.ColumnName);

                            Cell cell = new Cell();
                            cell.DataType = CellValues.String;
                            cell.CellValue = new CellValue(column.ColumnName);
                            headerRow.AppendChild(cell);
                        }
                        sheetData.AppendChild(headerRow);

                        //Add cells to sheetData
                        foreach (DataRow row in table.Rows)
                        {
                            Row newRow = new Row();
                            columns.ForEach(col =>
                            {
                                Cell cell = new Cell();
                                //If value is DBNull, do not set value to cell
                                if (row[col] != System.DBNull.Value)
                                {
                                    cell.DataType = CellValues.String;
                                    cell.CellValue = new CellValue(row[col].ToString());
                                }
                                newRow.AppendChild(cell);
                            });
                            sheetData.AppendChild(newRow);
                        }
                        result = $"Export {table.Rows.Count} rows of data to excel successfully.";
                    }
                }

                // Write the excel to Azure storage container
                using (FileStream fileStream = File.Open(filePath, FileMode.Open))
                {
                    bool exists = container.CreateIfNotExists();
                    var blob = container.GetBlockBlobReference(excelName);
                    blob.DeleteIfExists();
                    blob.UploadFromStream(fileStream);
                }
            }
            catch (Exception ex)
            {
                result =$"Export action failed. Error Message: {ex.Message}";
            }
            return result;
        }

        public static string ExcelImportToDB()
        {
            string result = string.Empty;
            try
            {
                //Check directory
                if (!Directory.Exists(directoryPath))
                {
                    Directory.CreateDirectory(directoryPath);
                }
                // Delete the file if it exists
                string filePath = $"{directoryPath}//{excelName}";
                if (File.Exists(filePath))
                {
                    File.Delete(filePath);
                }
                // Download blob to server disk.
                container.CreateIfNotExists();
                CloudBlockBlob blob = container.GetBlockBlobReference(excelName);
                blob.DownloadToFile(filePath, FileMode.Create);

                DataTable dt = new DataTable();
                using (SpreadsheetDocument spreadSheetDocument = SpreadsheetDocument.Open(filePath, false))
                {
                    //Get sheet data
                    WorkbookPart workbookPart = spreadSheetDocument.WorkbookPart;
                    IEnumerable<Sheet> sheets = spreadSheetDocument.WorkbookPart.Workbook.GetFirstChild<Sheets>().Elements<Sheet>();
                    string relationshipId = sheets.First().Id.Value;
                    WorksheetPart worksheetPart = (WorksheetPart)spreadSheetDocument.WorkbookPart.GetPartById(relationshipId);
                    Worksheet workSheet = worksheetPart.Worksheet;
                    SheetData sheetData = workSheet.GetFirstChild<SheetData>();
                    IEnumerable<Row> rows = sheetData.Descendants<Row>();

                    // Set columns
                    foreach (Cell cell in rows.ElementAt(0))
                    {
                        dt.Columns.Add(cell.CellValue.InnerXml);
                    }

                    //Write data to datatable
                    foreach (Row row in rows.Skip(1))
                    {
                        DataRow newRow = dt.NewRow();
                        for (int i = 0; i < row.Descendants<Cell>().Count(); i++)
                        {
                            if (row.Descendants<Cell>().ElementAt(i).CellValue != null)
                            {
                                newRow[i] = row.Descendants<Cell>().ElementAt(i).CellValue.InnerXml;
                            }
                            else
                            {
                                newRow[i] = DBNull.Value;
                            }
                        }
                        dt.Rows.Add(newRow);
                    }
                }

                //Bulk copy datatable to DB
                SqlBulkCopy bulkCopy = new SqlBulkCopy(connectionStr);
                try
                {
                    columns.ForEach(col => { bulkCopy.ColumnMappings.Add(col, col); });
                    bulkCopy.DestinationTableName = tableName;
                    bulkCopy.WriteToServer(dt);
                }
                catch (Exception ex)
                {
                    throw ex;
                }
                finally
                {
                    bulkCopy.Close();
                }
                result = $"Import {dt.Rows.Count} rows of data to DB successfully.";
            }
            catch (Exception ex)
            {
                result = $"Import action failed. Error Message: {ex.Message}";
            }
            return result;
        }

    }
}