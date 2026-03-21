using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

class PdfToXpsConverter
{
    static void Main()
    {
        // Create temporary folder for demo files.
        string tempFolder = Path.Combine(Path.GetTempPath(), "PdfToXpsDemo");
        Directory.CreateDirectory(tempFolder);

        // Paths for the source PDF and the resulting XPS files.
        string pdfFilePath = Path.Combine(tempFolder, "SourceDocument.pdf");
        string xpsFilePath = Path.Combine(tempFolder, "ConvertedDocument.xps");

        // -----------------------------------------------------------------
        // Step 1: Generate a simple PDF document (simulating an existing PDF).
        // -----------------------------------------------------------------
        Document tempDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(tempDoc);
        builder.Writeln("This is a sample PDF document created for conversion demo.");
        // Add a simple annotation (comment) to demonstrate preservation.
        Comment comment = new Comment(tempDoc, "Author", "Initials", DateTime.Now);
        comment.Paragraphs.Add(new Paragraph(tempDoc));
        comment.Paragraphs[0].AppendChild(new Run(tempDoc, "Sample annotation"));
        builder.CurrentParagraph.AppendChild(comment);
        tempDoc.Save(pdfFilePath, SaveFormat.Pdf);

        // -----------------------------------------------------------------
        // Step 2: Load the generated PDF and convert it to XPS while preserving annotations.
        // -----------------------------------------------------------------
        Document pdfDocument = new Document(pdfFilePath);
        XpsSaveOptions xpsOptions = new XpsSaveOptions
        {
            // Annotations are preserved by default; no extra settings required.
        };
        pdfDocument.Save(xpsFilePath, xpsOptions);

        Console.WriteLine($"PDF successfully converted to XPS.");
        Console.WriteLine($"Source PDF: {pdfFilePath}");
        Console.WriteLine($"Result XPS: {xpsFilePath}");
    }
}
