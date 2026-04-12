using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a temporary working folder.
        string workFolder = Path.Combine(Path.GetTempPath(), "AsposeDemo");
        Directory.CreateDirectory(workFolder);

        // Paths for the sample DOCX and the resulting PDF.
        string docxPath = Path.Combine(workFolder, "LargeDocument.docx");
        string pdfPath = Path.Combine(workFolder, "LargeDocument.pdf");

        // -----------------------------------------------------------------
        // 1. Generate a large DOCX document.
        // -----------------------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add a lot of text to make the document sizable.
        for (int i = 0; i < 5000; i++)
        {
            builder.Writeln($"This is line {i + 1} of a large document generated for streaming conversion.");
        }

        // Save the DOCX to disk (required before loading via stream).
        doc.Save(docxPath, SaveFormat.Docx);

        // -----------------------------------------------------------------
        // 2. Load the DOCX from a stream.
        // -----------------------------------------------------------------
        using (FileStream inputStream = File.OpenRead(docxPath))
        {
            // Load the document directly from the input stream.
            Document largeDoc = new Document(inputStream);

            // -----------------------------------------------------------------
            // 3. Prepare PDF save options with memory optimization.
            // -----------------------------------------------------------------
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                // Reduces memory consumption during saving at the cost of speed.
                MemoryOptimization = true
            };

            // -----------------------------------------------------------------
            // 4. Save the document to a MemoryStream (streaming conversion).
            // -----------------------------------------------------------------
            using (MemoryStream pdfStream = new MemoryStream())
            {
                largeDoc.Save(pdfStream, pdfOptions);

                // Ensure the stream contains data.
                if (pdfStream.Length == 0)
                {
                    throw new InvalidOperationException("PDF conversion produced an empty stream.");
                }

                // Reset position before reading or copying.
                pdfStream.Position = 0;

                // Optionally write the PDF to a file for verification.
                using (FileStream fileOut = File.Create(pdfPath))
                {
                    pdfStream.CopyTo(fileOut);
                }
            }
        }

        // Verify that the output PDF file exists and is not empty.
        if (!File.Exists(pdfPath) || new FileInfo(pdfPath).Length == 0)
        {
            throw new InvalidOperationException("PDF file was not created correctly.");
        }

        // Clean up temporary files (optional).
        // File.Delete(docxPath);
        // File.Delete(pdfPath);
        // Directory.Delete(workFolder, true);
    }
}
