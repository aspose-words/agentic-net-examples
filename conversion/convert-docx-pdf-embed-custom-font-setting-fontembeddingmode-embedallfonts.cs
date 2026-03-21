using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        Document doc;

        const string inputPath = "Input.docx";
        const string outputPath = "Output.pdf";

        if (File.Exists(inputPath))
        {
            // Load the source DOCX file.
            doc = new Document(inputPath);
        }
        else
        {
            // Create a simple document if the input file is missing.
            doc = new Document();
            var builder = new DocumentBuilder(doc);
            builder.Writeln("This is a sample document generated because Input.docx was not found.");
        }

        // Remove all fields (e.g., barcode fields) that could cause rendering issues.
        doc.Range.Fields.Clear();

        // Create PDF save options and configure font embedding.
        var pdfOptions = new PdfSaveOptions
        {
            // Embed all fonts (including custom fonts) into the PDF.
            FontEmbeddingMode = PdfFontEmbeddingMode.EmbedAll
        };

        // Save the document as PDF with the specified options.
        doc.Save(outputPath, pdfOptions);

        Console.WriteLine($"Document saved to '{outputPath}'.");
    }
}
