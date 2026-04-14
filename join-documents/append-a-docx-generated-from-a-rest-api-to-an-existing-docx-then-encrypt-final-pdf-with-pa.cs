using System;
using System.IO;
using System.Threading.Tasks;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Words.Loading;

namespace AsposeWordsJoinAndEncrypt
{
    public class Program
    {
        private const string ExistingDocFileName = "Existing.docx";
        private const string MergedPdfFileName = "MergedEncrypted.pdf";
        private const string PdfPassword = "SecretPassword";

        public static async Task Main()
        {
            // Prepare file paths in the current working directory.
            string baseDir = Directory.GetCurrentDirectory();
            string existingDocPath = Path.Combine(baseDir, ExistingDocFileName);
            string mergedPdfPath = Path.Combine(baseDir, MergedPdfFileName);

            // 1. Create a sample existing DOCX document.
            CreateSampleExistingDocument(existingDocPath);

            // 2. Simulate obtaining a DOCX from a REST API.
            using MemoryStream apiDocStream = await GetDocxFromRestApiAsync();

            // 3. Load the existing document from file.
            Document existingDoc = new Document(existingDocPath);

            // 4. Load the API‑generated document from the memory stream.
            Document apiDoc = new Document(apiDocStream);

            // 5. Append the API document to the existing one, preserving source formatting.
            existingDoc.AppendDocument(apiDoc, ImportFormatMode.KeepSourceFormatting);

            // 6. Save the merged document as an encrypted PDF.
            SaveMergedDocumentAsEncryptedPdf(existingDoc, mergedPdfPath, PdfPassword);

            // 7. Validate that the PDF was created and contains content from both sources.
            ValidateMergedPdf(mergedPdfPath, PdfPassword);

            // Indicate successful completion.
            Console.WriteLine("Document merged and encrypted PDF saved successfully.");
        }

        private static void CreateSampleExistingDocument(string filePath)
        {
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.Writeln("Existing document content.");
            doc.Save(filePath, SaveFormat.Docx);
        }

        // Simulates a REST API call that returns a DOCX file.
        private static async Task<MemoryStream> GetDocxFromRestApiAsync()
        {
            // In a real scenario you would use HttpClient to download the DOCX.
            // Here we generate it locally to keep the example self‑contained.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.Writeln("Content from REST API.");

            MemoryStream stream = new MemoryStream();
            doc.Save(stream, SaveFormat.Docx);
            stream.Position = 0; // Reset for reading.
            await Task.CompletedTask; // Mimic async behavior.
            return stream;
        }

        private static void SaveMergedDocumentAsEncryptedPdf(Document doc, string outputPath, string password)
        {
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                EncryptionDetails = new PdfEncryptionDetails(password, string.Empty)
            };
            doc.Save(outputPath, pdfOptions);
        }

        private static void ValidateMergedPdf(string pdfPath, string password)
        {
            if (!File.Exists(pdfPath))
                throw new FileNotFoundException("The encrypted PDF file was not created.", pdfPath);

            // Load the encrypted PDF using the password.
            LoadOptions loadOptions = new LoadOptions(password);
            Document pdfDoc = new Document(pdfPath, loadOptions);
            string text = pdfDoc.GetText();

            if (!text.Contains("Existing document content.") || !text.Contains("Content from REST API."))
                throw new InvalidOperationException("The merged PDF does not contain expected content.");
        }
    }
}
