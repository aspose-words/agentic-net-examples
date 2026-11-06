using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class PdfToHtmlConverter
{
    public static void Main()
    {
        // Create a sample PDF document.
        Document sampleDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sampleDoc);
        builder.Writeln("This is a sample PDF document with some text.");
        string pdfPath = "sample.pdf";
        sampleDoc.Save(pdfPath, SaveFormat.Pdf);

        if (!File.Exists(pdfPath) || new FileInfo(pdfPath).Length == 0)
            throw new InvalidOperationException("Failed to create the sample PDF file.");

        // Load the PDF document.
        Document pdfDocument = new Document(pdfPath);

        // Configure HTML save options.
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html);
        // Note: ExportEmbeddedFonts and ExportEmbeddedImages properties are not available
        // in the version of Aspose.Words used for this example. The default behavior will be applied.

        // Save the PDF as HTML.
        string htmlPath = "output.html";
        pdfDocument.Save(htmlPath, htmlOptions);

        // Validate the HTML output.
        if (!File.Exists(htmlPath) || new FileInfo(htmlPath).Length == 0)
            throw new InvalidOperationException("HTML conversion failed; output file was not created.");

        // Optionally, list any external image files that were generated.
        string imagesFolder = Path.Combine(
            Path.GetDirectoryName(htmlPath) ?? string.Empty,
            Path.GetFileNameWithoutExtension(htmlPath) + "_files");

        if (Directory.Exists(imagesFolder))
        {
            string[] imageFiles = Directory.GetFiles(imagesFolder);
            Console.WriteLine($"External images saved to folder: {imagesFolder}");
            foreach (string imageFile in imageFiles)
            {
                Console.WriteLine($" - {Path.GetFileName(imageFile)}");
            }
        }
        else
        {
            Console.WriteLine("No external image folder was created.");
        }

        Console.WriteLine("PDF successfully converted to HTML.");
    }
}
