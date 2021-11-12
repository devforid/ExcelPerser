using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Dynamic;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using ExcelDataReader;

namespace Excel_parse.Services
{
    public class ExcelParserService
    {
        private static string testUrl = "https://meta-test.swica.ch/cp/Home";
        private static string prodUrl = "https://meta.swica.ch/cp/Home";
        ExcelProcessingService excelProcessingService = new ExcelProcessingService(prodUrl);
        public ExcelParserService()
        {
        }
        public void ParseExcelFile()
        {
            var fileDirectory = "G:\\excel2\\LerbExportmitCode.xlsx";
            int successFullyParsed = 0;
            int parsedFailed = 0;
            ArrayList excelDataRows = new ArrayList();
            try
            {
                System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
                var dataTable = ReadFile(fileDirectory);
                successFullyParsed = dataTable.Rows.Count;
                Console.WriteLine(successFullyParsed);
                for(var i = 0; i < dataTable.Rows.Count; i++)
                {
                    try
                    {
                       var row = ProcessEachRow(dataTable.Rows[i]);
                        excelDataRows.Add(row);
                    }
                    catch(Exception ex)
                    {

                    }
                }
                excelProcessingService.GenerateExcelReport(excelDataRows);
            }
            catch (Exception ex)
            {

            }
        }
        private DataTable ReadFile(string fileDirectory)
        {
            DataTable dataTable = null;
            if (!string.IsNullOrWhiteSpace(fileDirectory))
            {
                using (var client = new WebClient())
                {
                    var content = client.DownloadData(fileDirectory);
                    using (var stream = new MemoryStream(content))
                    {
                        IExcelDataReader reader;
                        reader = ExcelDataReader.ExcelReaderFactory.CreateReader(stream);
                        var conf = new ExcelDataSetConfiguration
                        {
                            ConfigureDataTable = _ => new ExcelDataTableConfiguration
                            {
                                UseHeaderRow = true
                            }
                        };
                        var dataSet = reader.AsDataSet(conf);
                        dataTable = dataSet.Tables[1];
                    }
                }
            }

            return dataTable;
        }
        private dynamic ProcessEachRow(DataRow dataRow)
        {
            var itemArray = dataRow.ItemArray.ToArray();
            dynamic rowData = new ExpandoObject();
            rowData.Nummer = DBNull.Value != itemArray[0] ? Convert.ToString( itemArray[0]): null;
            rowData.UniqueCode = DBNull.Value != itemArray[1] ? Convert.ToString(itemArray[1]) : null;  
            rowData.Nachname = DBNull.Value != itemArray[2] ? Convert.ToString(itemArray[2]) : null;
            rowData.Vorname = DBNull.Value != itemArray[3] ? Convert.ToString(itemArray[3]) : null;
            rowData.Strasse = DBNull.Value != itemArray[4] ? Convert.ToString(itemArray[4]) : null;
            rowData.Postfach = DBNull.Value != itemArray[5] ? Convert.ToString(itemArray[5]) : null;
            rowData.PLZ = DBNull.Value != itemArray[6] ? Convert.ToString(itemArray[6]) : null;
            rowData.Ort = DBNull.Value != itemArray[7] ? Convert.ToString(itemArray[7]) : null;
            rowData.GebDat = DBNull.Value != itemArray[8] ? Convert.ToString(itemArray[8]) : null;
            rowData.Sprache = DBNull.Value != itemArray[9] ? Convert.ToString(itemArray[9]) : null;
            rowData.Mobile = DBNull.Value != itemArray[10] ? Convert.ToString(itemArray[10]) : null;
            rowData.Partnergruppe = DBNull.Value != itemArray[11] ? Convert.ToString(itemArray[11]) : null;
            rowData.Massnahme = DBNull.Value != itemArray[12] ? Convert.ToString(itemArray[12]) : null;
            rowData.MassnahmeVerwendungsart = DBNull.Value != itemArray[13] ? Convert.ToString(itemArray[13]) : null;
            return rowData;
        }
    }
}
