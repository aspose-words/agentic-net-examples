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
            // -----------------------------------------------------------------
            // Create a simple template document in memory.
            // -----------------------------------------------------------------
            Document doc = new Document();
            DocumentBuilder templateBuilder = new DocumentBuilder(doc);
            templateBuilder.Writeln("Sample Report");
            // (If the template needed merge fields they could be added here.)

            // -----------------------------------------------------------------
            // Prepare JSON data containing the attachment (Base64 + file name).
            // -----------------------------------------------------------------
            string attachmentContent = "Hello, this is the embedded file content.";
            string attachmentBase64 = Convert.ToBase64String(Encoding.UTF8.GetBytes(attachmentContent));
            string attachmentFileName = "Sample.txt";

            string jsonString = $@"{{
                ""AttachmentBase64"": ""{attachmentBase64}"",
                ""AttachmentFileName"": ""{attachmentFileName}""
            }}";

            // Load JSON data from a memory stream.
            using var jsonStream = new MemoryStream(Encoding.UTF8.GetBytes(jsonString));
            JsonDataSource jsonDataSource = new JsonDataSource(jsonStream);

            // Build the report (no specific data fields are used in this simple example).
            ReportingEngine engine = new ReportingEngine();
            engine.BuildReport(doc, jsonDataSource, "data");

            // -----------------------------------------------------------------
            // Extract attachment bytes from the JSON we just created.
            // -----------------------------------------------------------------
            byte[] attachmentBytes;
            using (JsonDocument jsonDoc = JsonDocument.Parse(jsonString))
            {
                JsonElement root = jsonDoc.RootElement;
                if (root.TryGetProperty("AttachmentBase64", out JsonElement base64Element) &&
                    root.TryGetProperty("AttachmentFileName", out JsonElement nameElement))
                {
                    attachmentBytes = Convert.FromBase64String(base64Element.GetString()!);
                }
                else
                {
                    throw new InvalidOperationException("Attachment data not found in JSON.");
                }
            }

            // Insert the attachment as an OLE object into the document.
            DocumentBuilder builder = new DocumentBuilder(doc);
            using (MemoryStream attachmentStream = new MemoryStream(attachmentBytes))
            {
                // "Package" progID allows embedding arbitrary files.
                builder.InsertOleObject(attachmentStream, "Package", false, null);
            }

            // -----------------------------------------------------------------
            // Save the final document as PDF with the attachment embedded.
            // -----------------------------------------------------------------
            string outputPdfPath = Path.Combine(Directory.GetCurrentDirectory(), "GeneratedReport.pdf");
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                AttachmentsEmbeddingMode = PdfAttachmentsEmbeddingMode.DocumentEmbeddedFiles
            };
            doc.Save(outputPdfPath, pdfOptions);

            Console.WriteLine($"PDF generated successfully: {outputPdfPath}");
        }
    }
}
