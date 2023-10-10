using GroupDocs.Parser;
using GroupDocs.Parser.Options;
using iText.Kernel.Pdf;
using iText.Kernel.Pdf.Canvas.Parser;
using iText.Kernel.Pdf.Canvas.Parser.Listener;
using Microsoft.AspNetCore.Mvc;
using System.Text;
using Spire.Doc;

using IHostingEnvironment = Microsoft.AspNetCore.Hosting.IHostingEnvironment;
using PdfDocument = iText.Kernel.Pdf.PdfDocument;
using Task = System.Threading.Tasks.Task;
using System.Globalization;
using CsvHelper.Configuration;
using CsvHelper;
using FileFormat = Spire.Doc.FileFormat;

namespace Intro_To_Textract.Controllers
{
    public class Textract : Controller
    {
        [Obsolete]
        private IHostingEnvironment Environment;

        [Obsolete]
        public Textract(IHostingEnvironment _environment)
        {
            Environment = _environment;
        }

        public IActionResult Index()
        {
            return View();
        }

        public IActionResult UploadFile()
        {
            return View();
        }

        [HttpPost]
        [Obsolete]
        public async Task<ActionResult> UploadFileAsync(IFormFile file)
        {
            string extractedData = await SaveFile(file);
            ViewBag.Message = extractedData;
            return View();
            // return RedirectToAction("ViewFile", new { objectKey = extractedData });

        }

        [Obsolete]
        public async Task<string> SaveFile(IFormFile file)
        {
            if (file == null)
            {
                return "Please select document for extract!";
            }

            string text = string.Empty;
            string path = Path.Combine(Environment.WebRootPath, "Uploads");
            if (!Directory.Exists(path))
            {
                Directory.CreateDirectory(path);
            }
            string fileName = Path.GetFileName(file.FileName);
            string fullPath = Path.Combine(path, fileName);
            using (FileStream stream = new FileStream(Path.Combine(path, fileName), FileMode.Create))
            {
                file.CopyTo(stream);
            }

            try
            {
                string fileExtention = Path.GetExtension(file.FileName).ToLower().Replace(".", "");
                switch (fileExtention)
                {
                    case "pdf":
                        Console.WriteLine("Data extraction from PDF file");
                        text = await GetTextFromPDF(fullPath);
                        break;
                    case "doc" or "docx":
                        Console.WriteLine("Data extraction from Word file");
                        text = await GetTextFromWordSpireDoc(fullPath);
                        break;
                    case "txt":
                        Console.WriteLine("Data extraction from Text file");
                        text = await GetTextFromTextDoc(fullPath);
                        break;
                    case "csv":
                        Console.WriteLine("Data extraction from csv file");
                        text = await GetTextFromCsv(fullPath);
                        break;


                }

            }
            catch (Exception ex)
            {
                string errorMessage = ex.Message;
            }

            return text;
        }

        private async Task<string> GetTextFromWord(string fullPath)
        {
            //FileStream fstream = new FileStream(fullPath, FileMode.Open, FileAccess.Read);
            //StreamReader sreader = new StreamReader(fstream);

            //var text = sreader.ReadToEnd();
            String text = string.Empty;
            // Create an instance of Parser class
            using (Parser parser = new Parser(fullPath))
            {
                // Check if the document supports formatted text extraction
                if (!parser.Features.FormattedText)
                {
                    Console.WriteLine("Document isn't supports formatted text extraction.");
                    return "Document isn't supports formatted text extraction.";
                }

                // Get the document info
                IDocumentInfo documentInfo = parser.GetDocumentInfo();
                // Check if the document has pages
                if (documentInfo.PageCount == 0)
                {
                    Console.WriteLine("Document hasn't pages.");
                    return "Document hasn't pages.";
                }

                // Iterate over pages
                for (int p = 0; p < documentInfo.PageCount; p++)
                {
                    // Print a page number 
                    Console.WriteLine(string.Format("Page {0}/{1}", p + 1, documentInfo.PageCount));
                    // Extract a formatted text into the reader
                    using (TextReader reader = parser.GetFormattedText(p, new FormattedTextOptions(FormattedTextMode.PlainText)))
                    {
                        // Print a formatted text from the document
                        text = await reader.ReadToEndAsync();

                    }
                }

                return text.ToString();
            }
        }

        private async Task<string> GetTextFromWordSpireDoc(string fullPath)
        {
            //Create and Load Document instance
            Document mydoc = new Document();
            mydoc.LoadFromFile(fullPath);
            //Save to HTML
            mydoc.SaveToFile("test.html", FileFormat.Html);
            string text = System.IO.File.ReadAllText("test.html");
            //File.WriteAllText("Extract.txt", text.ToString());
            return await Task.FromResult(text);

        }
        private async Task<string> GetTextFromPDF(string path)
        {
            StringBuilder text = new StringBuilder();
            PdfReader itextreader = new PdfReader(path);
            PdfDocument pdfdoc = new PdfDocument(itextreader);
            int numberofpages = pdfdoc.GetNumberOfPages();
            for (int page = 1; page <= numberofpages; page++)
            {
                ITextExtractionStrategy strategy = new SimpleTextExtractionStrategy();
                string currenttext = PdfTextExtractor.GetTextFromPage(pdfdoc.GetPage(page), strategy);
                currenttext = Encoding.UTF8.GetString(ASCIIEncoding.Convert(
                    Encoding.Default, Encoding.UTF8, Encoding.Default.GetBytes(currenttext)));
                text.Append(currenttext);
            }
            return await Task.FromResult(text.ToString());
        }

        private async Task<string> GetTextFromTextDoc(string path)
        {
            string text = System.IO.File.ReadAllText(path);
            return await Task.FromResult(text.ToString());
        }

        private async Task<string> GetTextFromCsv(string path)
        {
            StringBuilder sb = new StringBuilder();
            using (StreamReader streamReader = new StreamReader(path, System.Text.Encoding.UTF8, true))
            {
                var csvConfig = new CsvConfiguration(CultureInfo.CurrentCulture)
                {
                    HasHeaderRecord = true,
                    Comment = '#',
                    AllowComments = true,
                    Delimiter = ";",
                };

                using var csvReader = new CsvReader(streamReader, csvConfig);

                string? value = null;
                //sb.Append("<table class='table table-striped'>");
                while (csvReader.Read())
                {
                    //if (csvReader.ReadHeader())
                    //{
                    //    sb.Append("<thead>");
                    //}
                    //sb.Append("<tr>");
                    for (int i = 0; csvReader.TryGetField<string>(i, out value); i++)
                    {
                        Console.Write($"{value} ");
                        sb.Append(value);
                        //sb.Append("<td>"+ value + "</td>");
                    }
                    // sb.Append("</tr>");
                    Console.WriteLine();
                }

                // var result = JsonConvert.SerializeObject(listObjResult);
                ////csvData = JArray.Parse(result).ToString();
                //Console.WriteLine(csvData);
            }
            return await Task.FromResult(sb.ToString());
        }


    }
}
