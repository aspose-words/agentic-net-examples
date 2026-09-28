using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Prepare input and output folders.
        string inputFolder = "input";
        string outputFolder = "output";

        Directory.CreateDirectory(inputFolder);
        Directory.CreateDirectory(outputFolder);

        // Create sample PDF files.
        for (int i = 1; i <= 3; i++)
        {
            Document sampleDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(sampleDoc);
            builder.Writeln($"Sample PDF document #{i}");
            builder.Writeln("This is a paragraph with some text to demonstrate layout and font preservation.");

            string pdfPath = Path.Combine(inputFolder, $"sample{i}.pdf");
            sampleDoc.Save(pdfPath, SaveFormat.Pdf);
        }

        // Batch convert PDFs to HTML preserving layout and fonts.
        string[] pdfFiles = Directory.GetFiles(inputFolder, "*.pdf");
        foreach (string pdfFile in pdfFiles)
        {
            Document pdfDoc = new Document(pdfFile);

            HtmlSaveOptions htmlOptions = new HtmlSaveOptions
            {
                ExportFontResources = true,
                ExportPageMargins = true
                // ExportHeadersFootersMode is optional; omitted for compatibility with older library versions.
            };

            string htmlFileName = Path.GetFileNameWithoutExtension(pdfFile) + ".html";
            string htmlPath = Path.Combine(outputFolder, htmlFileName);
            pdfDoc.Save(htmlPath, htmlOptions);

            if (!File.Exists(htmlPath))
                throw new InvalidOperationException($"Expected HTML output was not created: {htmlPath}");
        }

        // Indicate completion.
        Console.WriteLine("Batch PDF to HTML conversion completed successfully.");
    }
}
