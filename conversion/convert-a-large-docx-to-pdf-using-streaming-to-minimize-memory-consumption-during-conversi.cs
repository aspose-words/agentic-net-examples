using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a large DOCX document programmatically.
        Document source = new Document();
        DocumentBuilder builder = new DocumentBuilder(source);

        // Add many pages to simulate a large document.
        for (int i = 0; i < 1000; i++)
        {
            builder.Writeln($"This is line {i + 1} of a large document.");
            // Insert a page break every 50 lines to increase size.
            if ((i + 1) % 50 == 0)
                builder.InsertBreak(BreakType.PageBreak);
        }

        // Save the source document to a local DOCX file.
        const string docxPath = "large.docx";
        source.Save(docxPath, SaveFormat.Docx);

        // Load the DOCX document for conversion.
        Document doc = new Document(docxPath);

        // Prepare a memory stream for PDF output.
        using (MemoryStream pdfStream = new MemoryStream())
        {
            // Create PDF save options with memory optimization enabled.
            SaveOptions pdfOptions = SaveOptions.CreateSaveOptions(SaveFormat.Pdf);
            pdfOptions.MemoryOptimization = true;

            // Save the document to the memory stream as PDF.
            doc.Save(pdfStream, pdfOptions);

            // Verify that data was written to the stream.
            if (pdfStream.Length == 0)
                throw new InvalidOperationException("PDF conversion produced an empty stream.");

            // Write the PDF bytes to a file.
            const string pdfPath = "large.pdf";
            File.WriteAllBytes(pdfPath, pdfStream.ToArray());

            // Validate that the PDF file was created.
            if (!File.Exists(pdfPath))
                throw new InvalidOperationException("Expected output PDF file was not created.");
        }
    }
}
