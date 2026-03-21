using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        const string inputPath = "input.docx";
        const string outputPath = "output.pdf";

        Document doc;

        // Load the source DOCX file if it exists; otherwise create a simple document.
        if (File.Exists(inputPath))
        {
            doc = new Document(inputPath);
        }
        else
        {
            doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.Writeln("This is a sample document generated because 'input.docx' was not found.");
        }

        // Configure PDF save options to apply high compression to images.
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            ImageCompression = PdfImageCompression.Jpeg,
            JpegQuality = 30
        };

        try
        {
            // Save the document as PDF using the configured options.
            doc.Save(outputPath, pdfOptions);
            Console.WriteLine($"PDF saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred while saving the PDF:");
            Console.WriteLine(ex);
        }
    }
}
