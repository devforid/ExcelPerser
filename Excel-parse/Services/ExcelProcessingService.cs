using NPOI.SS.UserModel;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Dynamic;
using System.IO;
using System.Net.Http;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Xml;
using System.Xml.Serialization;
using System.Linq;
using QRCoder;
using System.Drawing;
using Selise.Ecap.StorageService.Queries;
using Selise.Ecap.StorageService.ClientService;

namespace Excel_parse.Services
{
    public class ExcelProcessingService
    {
        HttpRequestService httpRequestService = new HttpRequestService();
        private string requestUrl;
        PostDataProcessResultSet postDataProcessResultset;
        private const string FILENAME = "g:\\PostDataProcessResultSet.xml";

        public ExcelProcessingService(string url)
        {
            requestUrl = url;
            postDataProcessResultset = DeSerializeObject<PostDataProcessResultSet>(FILENAME);

            if(postDataProcessResultset == null)
            {
                postDataProcessResultset = new PostDataProcessResultSet();
            }

            if(postDataProcessResultset.PostDataProcessResultList==null)
            {
                postDataProcessResultset.PostDataProcessResultList = new List<PostDataProcessResult>();
            }
        }

        public ExcelProcessingService()
        {
        }

        public string AssemblyDirectory
        {
            get
            {
                string codeBase = Assembly.GetExecutingAssembly().CodeBase;
                UriBuilder uri = new UriBuilder(codeBase);
                string path = Uri.UnescapeDataString(uri.Path);
                return Path.GetDirectoryName(path);
            }
        }

        public void GenerateExcelReport(ArrayList excelDataRows)
        {
            try
            {
                var workbook = CreateWorkBook();
                if (workbook == null) return;
                var sheet = workbook.GetSheetAt(0);
                var creationHelper = workbook.GetCreationHelper();
                var currentRow = 1;
                try
                {                    
                    foreach (dynamic rowData in excelDataRows)
                    {                        
                        PostDataProcessResult postDataProcessResult = new PostDataProcessResult
                        {
                            RowNumber = currentRow,
                            ItemId = rowData.UniqueCode,
                            IsProcessed = false
                        };

                        //if (currentRow >= 16000)
                        //{
                        //    break;
                        //}

                        currentRow++;

                        if (currentRow <= 16000)
                        {
                            continue;
                        }

                        IRow row = sheet.CreateRow(currentRow-16000);
                        
                        ExcelRowCreator rowCreator = new ExcelRowCreator(row, creationHelper);
                        //currentRow++;


                        PostData postData = new PostData();
                        postData.ItemId = Convert.ToString(rowData.UniqueCode);
                        //postData.Uri = "https://meta-test.swica.ch/cp/Home?uniqueCode=" + Convert.ToString(rowData.UniqueCode);
                        postData.Uri = "https://meta.swica.ch/cp/Home?uniqueCode=" + Convert.ToString(rowData.UniqueCode);

                        var requesturl = "https://urlshortener.selise.biz/v2/URLShortenerService/Shortener/CreateShortUrl";

                        

                        if (postDataProcessResultset.PostDataProcessResultList.Any(x => x.ItemId == postData.ItemId && x.IsProcessed == true))
                        {
                            Console.WriteLine("Processing Local : ###  " + currentRow);
                            rowData.ShortUrl = "https://urlshortener.selise.biz/v2/" + postData.ItemId;
                            string qrCodeImage = GenerateQRCode(rowData.ShortUrl);
                            //SaveFileInStorage(postData.ItemId, qrCodeStream);
                            rowData.QrCodeImageBase64 = qrCodeImage;
                            CreateRow(rowData, rowCreator);
                        }
                        else
                        {

                            //Task<HttpResponseMessage> task = httpRequestService.PostAsJsonAsync(requesturl.ToString(), postData);


                            HttpResponseMessage response = httpRequestService.PostDataAsJsonAsync(requesturl.ToString(), postData);

                            if (response != null && response.IsSuccessStatusCode)
                            {
                                postDataProcessResult.IsProcessed = true;
                                Console.WriteLine("Processing : ###  " + currentRow);
                                rowData.ShortUrl = "https://urlshortener.selise.biz/v2/" + postData.ItemId;
                                var qrCodeImage = GenerateQRCode(rowData.ShortUrl);
                                rowData.QrCodeImageBase64 = qrCodeImage;

                                CreateRow(rowData, rowCreator);
                            }
                            postDataProcessResultset.PostDataProcessResultList.Add(postDataProcessResult);
                        }
                        

                        //Task.Wait();                

                    }

                }
                catch (Exception ex)
                {
                    
                }
                

                SerializeObject<PostDataProcessResultSet>(postDataProcessResultset, FILENAME);


                SaveWorkBook(workbook);

            }
            catch (Exception ex)
            {

            }

        }


        public T DeSerializeObject<T>(string fileName)
        {
            if (string.IsNullOrEmpty(fileName)) { return default(T); }

            T objectOut = default(T);

            try
            {
                XmlDocument xmlDocument = new XmlDocument();
                xmlDocument.Load(fileName);
                string xmlString = xmlDocument.OuterXml;

                using (StringReader read = new StringReader(xmlString))
                {
                    Type outType = typeof(T);

                    XmlSerializer serializer = new XmlSerializer(outType);
                    using (XmlReader reader = new XmlTextReader(read))
                    {
                        objectOut = (T)serializer.Deserialize(reader);
                    }
                }
            }
            catch (Exception ex)
            {
                //Log exception here
            }

            return objectOut;
        }
        public void SerializeObject<T>(T serializableObject, string fileName)
        {
            if (serializableObject == null) { return; }

            try
            {
                XmlDocument xmlDocument = new XmlDocument();
                XmlSerializer serializer = new XmlSerializer(serializableObject.GetType());
                using (MemoryStream stream = new MemoryStream())
                {
                    serializer.Serialize(stream, serializableObject);
                    stream.Position = 0;
                    xmlDocument.Load(stream);
                    xmlDocument.Save(fileName);
                }
            }
            catch (Exception ex)
            {
                //Log exception here
            }
        }

        private IWorkbook CreateWorkBook()
        {
            string templateFile = AssemblyDirectory + Path.DirectorySeparatorChar + "ExcelTemplates" + Path.DirectorySeparatorChar + "ShortLinkTemplate.xlsx";
            IWorkbook workbook = null;
            try
            {
                using (var fs = new FileStream(templateFile, FileMode.Open, FileAccess.Read))
                {
                    workbook = WorkbookFactory.Create(fs);
                }
            }catch(Exception ex)
            {

            }

            return workbook;
        }


        private void CreateRow( dynamic rowData, ExcelRowCreator rowCreator)
        {
            rowCreator.CreateCell(0, rowData.Nummer);
            rowCreator.CreateCell(1, rowData.UniqueCode);
            rowCreator.CreateCell(2, rowData.Nachname);
            rowCreator.CreateCell(3, rowData.Vorname);
            rowCreator.CreateCell(4, rowData.Strasse);
            rowCreator.CreateCell(5, rowData.Postfach);
            rowCreator.CreateCell(6, rowData.PLZ);
            rowCreator.CreateCell(7, rowData.Ort);
            rowCreator.CreateCell(8, rowData.GebDat);
            rowCreator.CreateCell(9, rowData.Sprache);
            rowCreator.CreateCell(10,rowData.Mobile);
            rowCreator.CreateCell(11,rowData.Partnergruppe);
            rowCreator.CreateCell(12,rowData.Massnahme);
            rowCreator.CreateCell(13, rowData.MassnahmeVerwendungsart);
            rowCreator.CreateCell(14, rowData.ShortUrl);
            rowCreator.CreateCell(15, rowData.QrCodeImageBase64.Substring(0,31332));
            rowCreator.CreateCell(16, rowData.QrCodeImageBase64.Substring(31332));
        }

        private string GenerateQRCode(string context)
        {
            QRCodeGenerator qRCodeGenerator = new QRCodeGenerator();
            QRCodeData qRCodeData = qRCodeGenerator.CreateQrCode(context, QRCodeGenerator.ECCLevel.Q);
            QRCode qRCode = new QRCode(qRCodeData);
            Bitmap qrCodeImage = qRCode.GetGraphic(20);

            using (MemoryStream stream = new MemoryStream())
            {
                using (var bitmap = new Bitmap(qrCodeImage))
                {
                    bitmap.Save(stream, System.Drawing.Imaging.ImageFormat.Jpeg);
                    string imageUrl = "data:image/png;base64," + Convert.ToBase64String(stream.GetBuffer()); //Get Base64
                    return imageUrl;
                }

                //qrCodeImage.Save(stream, System.Drawing.Imaging.ImageFormat.Png);
                //byte[] byteImage = stream.ToArray();
                //string imageUrl = "data:image/png;base64," + Convert.ToBase64String(byteImage);
                //return imageUrl;
            }
        }

        private MemoryStream Base64ToImage(string base64String)
        {
            byte[] imageBytes = Convert.FromBase64String(base64String);
            MemoryStream ms = new MemoryStream(imageBytes, 0, imageBytes.Length);
            ms.Write(imageBytes, 0, imageBytes.Length);
            //System.Drawing.Image image = System.Drawing.Image.FromStream(ms, true);
         
            return ms;
        }
        private string base64String(Bitmap qrCodeImage)
        {
            using (MemoryStream stream = new MemoryStream())
            {
                using (var bitmap = new Bitmap(qrCodeImage))
                {
                    bitmap.Save(stream, System.Drawing.Imaging.ImageFormat.Jpeg);
                    return Convert.ToBase64String(stream.GetBuffer()); ;
                }
            }
        }

        //public string SaveFileInStorage(string uniqueCode, Stream image)
        //{
        //    GetPreSignedUrlForUploadQuery query = new GetPreSignedUrlForUploadQuery
        //    {
        //        ItemId = uniqueCode,
        //        Name = uniqueCode+".png",
        //        ParentDirectoryId = "Swica-QR-Files",
        //        Tags = "[\"test\"]",
        //        MetaData = ""
        //    };
        //    var response = this._storageClientService.GetPreSignedUrlForUpload(query).Result;
        //    if (!string.IsNullOrEmpty(response.UploadUrl))
        //    {
        //        using (var client = new HttpClient())
        //        {
        //            var res = client.PutAsync(response.UploadUrl, new StreamContent(image)).Result
        //                           .EnsureSuccessStatusCode();
        //            if (res.IsSuccessStatusCode)
        //            {
        //                return res.ToString();
        //            }
        //            return null;
        //        }
        //    }
        //    return null;
        //}

        private void SaveWorkBook(IWorkbook workbook)
        {
            try
            {
                using (var stream = new MemoryStream())
                {                    
                    workbook.Write(stream);
                    workbook.Close();
                    stream.Position = 0;
                    using (var fileStream = File.Create("g:\\prod-ShortLinkCode3.xlsx"))
                    {
                        stream.Seek(0, SeekOrigin.Begin);
                        stream.CopyTo(fileStream);
                    }
                    stream.Close();
                }

            }
            catch(Exception Ex)
            {

            }
        }
        //private dynamic getShortLink(dynamic rowValue)
        //{
        //    if (rowValue == null) return null;
        //    dynamic postData = new ExpandoObject();
        //    postData.ItemId = Guid.NewGuid().ToString();
        //    postData.Uri = "https://meta-test.swica.ch/cp/Home?uniqueCode="+ rowValue;

        //    var requesturl = "https://stg-urlshortener.selise.biz/v6/URLShortenerService/Shortener/CreateShortUrl";
        //    HttpResponseMessage response = httpRequestService.PostAsJsonAsync(requesturl.ToString(), postData);
        //    return response;
        //}
    }
}
