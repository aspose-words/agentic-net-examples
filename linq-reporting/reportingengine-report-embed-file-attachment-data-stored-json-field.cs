using System;
using System.IO;
using System.Text;
using System.Text.Json;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Saving;

namespace AsposeWordsReportWithAttachment
{
    class Program
    {
        static void Main()
        {
            // Paths to the template document and the JSON data file.
            string templatePath = @"C:\Data\ReportTemplate.docx";
            string jsonPath = @"C:\Data\ReportData.json";
            string outputPdfPath = @"C:\Data\GeneratedReport.pdf";

            // Load the template document.
            Document doc = new Document(templatePath);

            // Create a JsonDataSource from the JSON file (default parsing options).
            JsonDataSource jsonDataSource = new JsonDataSource(jsonPath);

            // Build the report using the ReportingEngine.
            ReportingEngine engine = new ReportingEngine();
            // The data source name ("data") can be used inside the template if needed.
            engine.BuildReport(doc, jsonDataSource, "data");

            // -----------------------------------------------------------------
            // Extract the attachment data from the same JSON file.
            // The JSON is expected to contain a property "AttachmentBase64"
            // with the file content encoded in Base64.
            // -----------------------------------------------------------------
            byte[] attachmentBytes;
            string attachmentFileName;

            using (FileStream jsonStream = File.OpenRead(jsonPath))
            {
                using (JsonDocument jsonDoc = JsonDocument.Parse(jsonStream))
                {
                    JsonElement root = jsonDoc.RootElement;

                    // Example JSON structure:
                    // {
                    //   "AttachmentBase64": "UEsDBBQAAAAIA...",
                    //   "AttachmentFileName": "Sample.xlsx"
                    // }
                    if (root.TryGetProperty("AttachmentBase64", out JsonElement base64Element) &&
                        root.TryGetProperty("AttachmentFileName", out JsonElement nameElement))
                    {
                        string base64String = base64Element.GetString();
                        attachmentFileName = nameElement.GetString();

                        attachmentBytes = Convert.FromBase64String(base64String);
                    }
                    else
                    {
                        throw new InvalidOperationException("Attachment data not found in JSON.");
                    }
                }
            }

            // Insert the attachment as an OLE object into the document.
            // The "Package" progID allows embedding arbitrary files.
            DocumentBuilder builder = new DocumentBuilder(doc);
            using (MemoryStream attachmentStream = new MemoryStream(attachmentBytes))
            {
                // Insert the OLE object at the current cursor position.
                // Parameters: stream, progId, isLink (false = embed), iconStream (null = default icon), fileName (optional)
                builder.InsertOleObject(attachmentStream, "Package", false, null);
            }

            // Save the document as PDF with the attachment embedded as a document file.
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                AttachmentsEmbeddingMode = PdfAttachmentsEmbeddingMode.DocumentEmbeddedFiles
            };

            doc.Save(outputPdfPath, pdfOptions);
        }
    }
}
