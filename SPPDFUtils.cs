using Azure.Identity;
using DocumentFormat.OpenXml.Packaging;
using iTextSharp.text;
using iTextSharp.text.html.simpleparser;
using iTextSharp.text.pdf;
using iTextSharp.tool.xml;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Azure.Functions.Worker;
using Microsoft.Azure.Functions.Worker.Http;
using Microsoft.Extensions.Logging;
using Microsoft.Graph;
using Microsoft.Graph.Models;
using Microsoft.Graph.Models.ODataErrors;
using Microsoft.Identity.Client;
using OpenXmlPowerTools;
using PnP.Core.Services;
using System.Drawing.Imaging;
using System.IO;
using System.Text;
using System.Xml.Linq;
using DriveUpload = Microsoft.Graph.Drives.Item.Items.Item.CreateUploadSession;
using Rectangle = iTextSharp.text.Rectangle;


namespace ERLLC.Functions
{
    public class SPPDFUtils
    {
        private readonly ILogger<SPPDFUtils> _logger;

        private readonly IPnPContextFactory _pnpContextFactory;

        public SPPDFUtils(ILogger<SPPDFUtils> logger, IPnPContextFactory pnpContextFactory)
        {
            _logger = logger;
            _pnpContextFactory = pnpContextFactory;
        }

        [Function("SPPDFUtils")]
        public async Task<IActionResult> RunAsync([HttpTrigger(AuthorizationLevel.Anonymous, "get", "post")] HttpRequest req)
        {
            _logger.LogInformation("C# HTTP trigger function processed a request.");
            string strPDFName = string.Empty;
            // Get request body
            var dataInput = await req.ReadFromJsonAsync<JoinPDFRequest>();
            try
            {
                var graphServiceClient = getGraphClient();
                ////Getting current execution path for creating temp files
                string tempPath = System.Reflection.Assembly.GetExecutingAssembly().Location;
                var directory = System.IO.Path.GetDirectoryName(tempPath);
                var siteDrives = await graphServiceClient.Sites[dataInput.siteId].Drives.GetAsync();
                var driveItem = siteDrives.Value.Where(d => d.Name == dataInput.pdfLibrary).FirstOrDefault();
                var lstFileBytes = GetFileBytes(graphServiceClient, dataInput.files, _logger).Result;
                if (lstFileBytes != null && lstFileBytes.Count > 0)
                {
                    var lstPDFBytes = new List<byte[]>();
                    for (int i = 0; i < dataInput.files.Length; i++)
                    {
                        if (dataInput.files[i].ToLower().EndsWith(".docx"))
                        {
                            lstPDFBytes.Add(ConvertWordToPDF(lstFileBytes[i], directory, _logger));
                        }
                        else if (dataInput.files[i].ToLower().EndsWith(".html"))
                        {
                            lstPDFBytes.Add(ConvertHTMLToPDF(lstFileBytes[i], _logger));
                        }
                        else if (dataInput.files[i].ToLower().EndsWith(".pdf"))
                        {
                            lstPDFBytes.Add(lstFileBytes[i]);
                        }
                        else if (dataInput.files[i].ToLower().EndsWith("png"))
                        {
                            lstPDFBytes.Add(ConvertImageToPDF(lstFileBytes[i], _logger));
                        }
                        else if (dataInput.files[i].ToLower().EndsWith("gif"))
                            lstPDFBytes.Add(ConvertImageToPDF(lstFileBytes[i], _logger));
                        else if (dataInput.files[i].ToLower().EndsWith("bmp"))
                            lstPDFBytes.Add(ConvertImageToPDF(lstFileBytes[i], _logger));
                        else if (dataInput.files[i].ToLower().EndsWith("jpeg"))
                            lstPDFBytes.Add(ConvertImageToPDF(lstFileBytes[i], _logger));
                        else if (dataInput.files[i].ToLower().EndsWith("tiff"))
                        {
                            lstPDFBytes.Add(ConvertImageToPDF(lstFileBytes[i], _logger));
                        }
                        else if (dataInput.files[i].ToLower().EndsWith("x-wmf"))
                        {
                            lstPDFBytes.Add(ConvertImageToPDF(lstFileBytes[i], _logger));
                        }
                    }
                    var finalPDFBytes = JoinPDF(lstPDFBytes, dataInput.strWaterMark, _logger);
                    strPDFName = "ERLLC_" + dataInput.id.ToString() + "_" + DateTime.Now.ToString("MMddyyyy_hhmmss") + ".pdf";
                    UploadFile(graphServiceClient, finalPDFBytes, driveItem.Id, strPDFName, _logger);

                }

            }
            catch (Exception ex)
            {
                _logger.LogInformation("Error:" + ex.Message.ToString());
            }


            return new OkObjectResult(dataInput.strSiteUrl + "/" + dataInput.pdfLibrary + "/" + strPDFName);
        }

        public async static Task<List<byte[]>> GetFileBytes(GraphServiceClient graphServiceClient, string[] strFiles, ILogger<SPPDFUtils> _logger)
        {
            List<byte[]> lstFileBytes = new List<byte[]>();
            try
            {
                foreach (var fileAbsoluteUrl in strFiles)
                {
                    string base64Value = System.Convert.ToBase64String(System.Text.Encoding.UTF8.GetBytes(fileAbsoluteUrl));
                    string encodedUrl = "u!" + base64Value.TrimEnd('=').Replace('/', '_').Replace('+', '-');
                    var driveItemStream = await graphServiceClient.Shares[encodedUrl].DriveItem.Content.GetAsync();
                    byte[] driveItemBytes;
                    using (var streamReader = new MemoryStream())
                    {
                        driveItemStream.CopyTo(streamReader);
                        driveItemBytes = streamReader.ToArray();
                        lstFileBytes.Add(driveItemBytes);
                    }
                }
            }
            catch (Exception)
            {

                throw;
            }
            return lstFileBytes;
        }

        public async static Task UploadFile(GraphServiceClient graphServiceClient, byte[] pdfBytes, string driveId, string strFileName, ILogger<SPPDFUtils> _logger)
        {
            try
            {
                Stream fileStream = new MemoryStream(pdfBytes);
                var uploadSessionRequestBody = new DriveUpload.CreateUploadSessionPostRequestBody
                {
                    Item = new DriveItemUploadableProperties
                    {
                        AdditionalData = new Dictionary<string, object>
                            {
                                { "@microsoft.graph.conflictBehavior", "replace" },
                            },
                    },
                };

                // Create the upload session
                // itemPath does not need to be a path to an existing item
                var uploadSession = await graphServiceClient.Drives[driveId]
                    //Drives["b!VypNVofBGU2_wXzyt1MpxT2iYQoHEjlFsH0B9quOrCCY82fXI8ySS4th5DDPHiyO"]
                    .Items["root"]
                    .ItemWithPath(strFileName)
                    .CreateUploadSession
                    .PostAsync(uploadSessionRequestBody);

                // Max slice size must be a multiple of 320 KiB
                int maxSliceSize = 320 * 1024;
                var fileUploadTask = new LargeFileUploadTask<DriveItem>(
                    uploadSession, fileStream, maxSliceSize, graphServiceClient.RequestAdapter);

                var totalLength = fileStream.Length;
                // Create a callback that is invoked after each slice is uploaded
                IProgress<long> progress = new Progress<long>(prog =>
                {
                    Console.WriteLine($"Uploaded {prog} bytes of {totalLength} bytes");
                });

                try
                {
                    // Upload the file
                    var uploadResult = await fileUploadTask.UploadAsync(progress);

                    Console.WriteLine(uploadResult.UploadSucceeded ?
                        $"Upload complete, item ID: {uploadResult.ItemResponse.Id}" :
                        "Upload failed");
                }
                catch (ODataError ex)
                {
                    Console.WriteLine($"Error uploading: {ex.Error?.Message}");
                }
            }
            catch (Exception)
            {

                throw;
            }
        }

        public static byte[] ConvertWordToPDF(byte[] wordBytes, string strTempDir, ILogger<SPPDFUtils> _logger)
        {
            byte[] bytes = null;
            Stream stream = null;
            try
            {
                //Creating HTML stream using word file content
                string strHTMLContent = string.Empty;
                using (MemoryStream memoryStream = new MemoryStream())
                {
                    memoryStream.Write(wordBytes, 0, wordBytes.Length);
                    using (WordprocessingDocument doc =
                        WordprocessingDocument.Open(memoryStream, true))
                    {
                        int imageCounter = 0;
                        //Mapping images to HTML file
                        HtmlConverterSettings settings = new HtmlConverterSettings()
                        {
                            PageTitle = "",
                            ImageHandler = imageInfo =>
                            {
                                ++imageCounter;
                                string extension = imageInfo.ContentType.Split('/')[1].ToLower();
                                ImageFormat imageFormat = null;
                                if (extension == "png")
                                {
                                    extension = "gif";
                                    imageFormat = ImageFormat.Gif;
                                }
                                else if (extension == "gif")
                                    imageFormat = ImageFormat.Gif;
                                else if (extension == "bmp")
                                    imageFormat = ImageFormat.Bmp;
                                else if (extension == "jpeg")
                                    imageFormat = ImageFormat.Jpeg;
                                else if (extension == "tiff")
                                {
                                    extension = "gif";
                                    imageFormat = ImageFormat.Gif;
                                }
                                else if (extension == "x-wmf")
                                {
                                    extension = "wmf";
                                    imageFormat = ImageFormat.Wmf;
                                }
                                if (imageFormat == null)
                                    return null;

                                //string imageFileName = strTempDir + "/img" + "/image" +
                                string strTimeImg = "TempImg_" + DateTime.Now.ToString("MMddyyyy_hhmmss");
                                string imageFileName = Path.Combine(strTempDir, strTimeImg) +
                                    imageCounter.ToString() + "." + extension;
                                try
                                {
                                    _logger.LogInformation("saving image" + imageFileName + ", Format:" + imageFormat);
                                    imageInfo.Bitmap.Save(imageFileName, imageFormat);
                                }
                                catch (System.Runtime.InteropServices.ExternalException)
                                {
                                    return null;
                                }
                                XElement img = new XElement(Xhtml.img,
                                    new XAttribute(NoNamespace.src, imageFileName),
                                    imageInfo.ImgStyleAttribute,
                                    imageInfo.AltText != null ?
                                        new XAttribute(NoNamespace.alt, imageInfo.AltText) : null);
                                return img;
                            }
                        };
                        XElement html = HtmlConverter.ConvertToHtml(doc, settings);
                        strHTMLContent = html.ToStringNewLineOnAttributes();
                    }
                }
                StringReader stringReader = new StringReader(strHTMLContent);
                Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
                using (MemoryStream memoryStream = new MemoryStream())
                {
                    //Converting HTML Stream to PDF using XMLWorkerhelper
                    iTextSharp.text.Document doc = new iTextSharp.text.Document();
                    PdfPTable tableLayout = new PdfPTable(4);
                    PdfWriter writer = PdfWriter.GetInstance(doc, memoryStream);
                    doc.Open();
                    XMLWorkerHelper.GetInstance().ParseXHtml(writer, doc, stringReader);
                    doc.Close();
                    stream = memoryStream;
                    bytes = memoryStream.ToArray();
                    memoryStream.Close();
                }

                DirectoryInfo taskDirectory = new DirectoryInfo(strTempDir);
                taskDirectory.GetFiles("TempImg_*").ToList().ForEach(file =>
                {
                    _logger.LogInformation("removing img: " + file.FullName);
                    file.Delete();
                    _logger.LogInformation("removed img: " + file.FullName);
                }
                );
            }
            catch (Exception ex)
            {
                //Adding log 
                _logger.LogInformation("Error in ConvertWordToPDF:" + ex.Message.ToString());
            }
            return bytes;
        }

        public static byte[] JoinPDF(List<byte[]> lstPdfBytes, string strWaterMark, ILogger<SPPDFUtils> _logger)
        {
            byte[] bytes = null;
            try
            {
                using(var streamPDFNew = new MemoryStream())
                {
                    //Joining pdf files to single file
                    iTextSharp.text.Document pdfDoc = new iTextSharp.text.Document(iTextSharp.text.PageSize.A4);
                    PdfCopy pdf = new PdfCopy(pdfDoc, streamPDFNew);
                    pdfDoc.Open();
                    pdf.AddDocument(new PdfReader(lstPdfBytes[0]));
                    for (int i = 1; i < lstPdfBytes.Count; i++)
                    {
                        pdf.AddDocument(new PdfReader(lstPdfBytes[i]));
                    }
                    if (pdfDoc != null)
                        pdfDoc.Close();
                    bytes= streamPDFNew.ToArray();

                }
                //Adding Page numbers and watermark to the combined PDF file
                bytes=AddPageNumbersWatermark(bytes, strWaterMark, _logger);

            }
            catch (Exception ex)
            {
                _logger.LogInformation("Error in JoinPDF:" + ex.Message.ToString() + ", InnerException:" + ex.InnerException != null && ex.InnerException.Message != null ? ex.InnerException.Message.ToString() : "");
            }
            return bytes;
        }

        public static byte[] ConvertHTMLToPDF(byte[] htmlBytes, ILogger<SPPDFUtils> _logger)
        {
            byte[] bytes = null;
            try
            {
                //Reading HTML file from SharePoint
                Stream fileStream = new MemoryStream(htmlBytes);
                StreamReader sReader = new StreamReader(fileStream);
                StringReader stringReader = new StringReader(RemoveHiddenTags(sReader.ReadToEnd(), _logger));
                Document pdfDoc = new Document(PageSize.A4, 10f, 10f, 10f, 0f);
                HTMLWorker htmlparser = new HTMLWorker(pdfDoc);
                using (MemoryStream memoryStream = new MemoryStream())
                {
                    PdfWriter writer = PdfWriter.GetInstance(pdfDoc, memoryStream);
                    pdfDoc.Open();

                    htmlparser.Parse(stringReader);
                    pdfDoc.Close();

                    bytes = memoryStream.ToArray();
                    memoryStream.Close();
                }
            }
            catch (Exception ex)
            {
                //Adding log 
                _logger.LogInformation("Error in ConvertHTMLToPDF:" + ex.Message.ToString());
            }
            return bytes;
        }

        public static string RemoveHiddenTags(string strHTMLStrig, ILogger<SPPDFUtils> _logger)
        {
            string strHTML = strHTMLStrig;
            try
            {
                //Replacing paragraph hidden tags
                while (strHTML.IndexOf("<p style='display:none'", StringComparison.InvariantCulture) > -1)
                {
                    //Replacing paragraph hidden tags
                    int iStartIndex = strHTML.IndexOf("<p style='display:none'", StringComparison.InvariantCulture);
                    int iEndIndex = strHTML.IndexOf("</p>", iStartIndex, StringComparison.InvariantCulture);
                    int itempStart = strHTML.IndexOf("<p", iStartIndex + 2, iEndIndex - iStartIndex, StringComparison.InvariantCulture);
                    int itempEnd = -1;
                    while (itempStart > -1)
                    {
                        itempEnd = iEndIndex;
                        iEndIndex = strHTML.IndexOf("</p>", iEndIndex + 2, StringComparison.InvariantCulture);
                        itempStart = strHTML.IndexOf("<p", itempEnd + 2, iEndIndex - itempEnd, StringComparison.InvariantCulture);
                    }
                    //iEndIndex = strHTML.IndexOf("</p>", iEndIndex + 2, StringComparison.InvariantCulture);
                    strHTML = strHTML.Remove(iStartIndex, iEndIndex - iStartIndex + 4);
                }
                //Replacing Div hidden tags
                strHTML = RemoveRecursiveTags(strHTML, _logger);
            }
            catch (Exception ex)
            {
                //Adding log 
                _logger.LogInformation("Error in RemoveHiddenTags:" + ex.Message.ToString());
            }
            return strHTML;
        }

        public static string RemoveRecursiveTags(string strHTML, ILogger<SPPDFUtils> _logger)
        {
            try
            {
                while (strHTML.IndexOf("<div style='display:none'", StringComparison.InvariantCulture) > -1)
                {
                    //Replacing paragraph hidden tags
                    int iStartIndex = strHTML.IndexOf("<div style='display:none'", StringComparison.InvariantCulture);
                    int iEndIndex = strHTML.IndexOf("</div>", iStartIndex, StringComparison.InvariantCulture);
                    int itempStart = strHTML.IndexOf("<div", iStartIndex + 2, iEndIndex - iStartIndex, StringComparison.InvariantCulture);
                    while (itempStart > -1)
                    {
                        if (strHTML.IndexOf("<div", itempStart + 2, iEndIndex - itempStart, StringComparison.InvariantCulture) > -1)
                        {
                            int innertempStart = strHTML.IndexOf("<div", itempStart + 2, iEndIndex - itempStart, StringComparison.InvariantCulture);
                            while (innertempStart > -1)
                            {
                                if (strHTML.IndexOf("<div", innertempStart + 2, iEndIndex - innertempStart, StringComparison.InvariantCulture) > -1)
                                {
                                    innertempStart = strHTML.IndexOf("<div", innertempStart + 2, iEndIndex - innertempStart, StringComparison.InvariantCulture);
                                }
                                else
                                {
                                    strHTML = strHTML.Remove(innertempStart, iEndIndex - innertempStart + 6);
                                    break;
                                }
                            }
                            iEndIndex = strHTML.IndexOf("</div>", iStartIndex, StringComparison.InvariantCulture);
                            itempStart = strHTML.IndexOf("<div", iStartIndex + 2, iEndIndex - iStartIndex, StringComparison.InvariantCulture);
                        }
                        else
                        {
                            strHTML = strHTML.Remove(itempStart, iEndIndex - itempStart + 6);
                            iEndIndex = strHTML.IndexOf("</div>", iStartIndex, StringComparison.InvariantCulture);
                            itempStart = strHTML.IndexOf("<div", iStartIndex + 2, iEndIndex - iStartIndex, StringComparison.InvariantCulture);
                        }
                    }
                    strHTML = strHTML.Remove(iStartIndex, iEndIndex - iStartIndex + 6);
                }
            }
            catch (Exception ex)
            {
                //Adding log 
                _logger.LogInformation("Error in RemoveHiddenTags:" + ex.Message.ToString());
            }
            return strHTML;
        }

        public static byte[] ConvertImageToPDF(byte[] imgBytes, ILogger<SPPDFUtils> _logger)
        {
            byte[] bytes = null;
            try
            {
                //Reading HTML file from SharePoint
                Stream fileStream = new MemoryStream(imgBytes);
                StreamReader sReader = new StreamReader(fileStream);
                Document pdfDoc = new Document(PageSize.A4, 10f, 10f, 10f, 0f);
                using (MemoryStream memoryStream = new MemoryStream())
                {
                    PdfWriter writer = PdfWriter.GetInstance(pdfDoc, memoryStream);
                    pdfDoc.Open();
                    iTextSharp.text.Image pic = iTextSharp.text.Image.GetInstance(fileStream);
                    pic.ScaleToFit(pdfDoc.PageSize);
                    pic.SetAbsolutePosition(0, 0);
                    pdfDoc.Add(pic);
                    pdfDoc.NewPage();
                    pdfDoc.Close();

                    bytes = memoryStream.ToArray();
                    memoryStream.Close();
                }
            }
            catch (Exception ex)
            {
                //Adding log 
                _logger.LogInformation("Error in ConvertImageToPDF:" + ex.Message.ToString());
            }
            return bytes;
        }

        public static byte[] AddPageNumbersWatermark(byte[] bytes, string strWatermark, ILogger<SPPDFUtils> _logger)
        {
            byte[] pdfbytes = null;
            try
            {
                    //System.IO.File.ReadAllBytes(strFile);
                iTextSharp.text.Font blackFont = FontFactory.GetFont("Arial", 12, iTextSharp.text.Font.NORMAL, BaseColor.BLACK);
                using (MemoryStream stream = new MemoryStream())
                {
                    PdfReader reader = new PdfReader(bytes);
                    using (PdfStamper stamper = new PdfStamper(reader, stream))
                    {
                        int pages = reader.NumberOfPages;
                        for (int i = 1; i <= pages; i++)
                        {
                            //Adding Page numbers
                            ColumnText.ShowTextAligned(stamper.GetUnderContent(i), Element.ALIGN_RIGHT,
                                                            new Phrase(i.ToString() + " of " + pages.ToString(), blackFont), 568f, 15f, 0);

                            //Adding Water marks
                            PdfContentByte pdfPageContents = stamper.GetUnderContent(i);
                            pdfPageContents.BeginText();
                            Rectangle pageSize = reader.GetPageSizeWithRotation(i);
                            BaseFont baseFont = BaseFont.CreateFont(BaseFont.HELVETICA_BOLD, Encoding.ASCII.EncodingName, false);
                            pdfPageContents.SetFontAndSize(baseFont, 40); // 40 point font
                            pdfPageContents.SetRGBColorFill(230, 228, 228);
                            double radians = Math.Atan2(pageSize.Height, pageSize.Width); // Get Radians for Atan2
                            double angle = radians * (180 / Math.PI);
                            float textAngle = (float)angle;
                            pdfPageContents.ShowTextAligned(PdfContentByte.ALIGN_CENTER, strWatermark, pageSize.Width / 2, pageSize.Height / 2, textAngle);
                            pdfPageContents.EndText();
                        }
                    }
                    pdfbytes = stream.ToArray();
                }
                //System.IO.File.WriteAllBytes(strFile, bytes);
            }
            catch (Exception ex)
            {
                //Adding log 
                _logger.LogInformation("Error in AddPageNumbersWatermark:" + ex.Message.ToString());
            }
            return pdfbytes;
        }

        public static GraphServiceClient getGraphClient()
        {
            //Host details
            //string strClientSecret = "Z9z8Q~xlxEqvWwREpS1lyfJgpG6gpeFYcBZetcbY";
            //string ClientId = "ee56e9ca-a43c-4973-8fec-e9b72d281da9";
            //var tenantId = "3b94ec52-97ed-4b61-8faa-d7a7596241f9";
            //Hunter Details
            string strClientSecret = "GBq8Q~vpkxqDdmONskq.TH~t2j48zSgCkmC8Ecf7";
            string ClientId = "14323ad7-61ac-4d6c-8cde-b57b932611ae";
            var tenantId = "2f146834-d37a-4dc7-a70f-19345d112002";
            IConfidentialClientApplication confidentialClientApplication = ConfidentialClientApplicationBuilder
           .Create(ClientId)
           .WithTenantId(tenantId)
           .WithClientSecret(strClientSecret)
           .Build();
            var clientSecretCredential = new ClientSecretCredential(tenantId, ClientId, strClientSecret);
            var scopes = new[] { "https://graph.microsoft.com/.default" };
            // Build the Microsoft Graph client
            GraphServiceClient graphServiceClient = new GraphServiceClient(clientSecretCredential, scopes);

            return graphServiceClient;
        }

        public class JoinPDFRequest
        {
            public string strSiteUrl { get; set; }
            public string siteId { get; set; }
            public string pdfLibrary { get; set; }
            public string[] files { get; set; }
            public string strWaterMark { get; set; }
            public int id { get; set; }
        }

    }
}
