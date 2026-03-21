using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Use files relative to the current working directory so they always exist.
        string inputPath = Path.Combine(Environment.CurrentDirectory, "Sample.docx");
        string outputPath = Path.Combine(Environment.CurrentDirectory, "Sample.mht");

        Document doc;

        if (File.Exists(inputPath))
        {
            // Load the existing DOCX file.
            doc = new Document(inputPath);
        }
        else
        {
            // Create a simple document if the input file is missing.
            doc = new Document();
            var builder = new DocumentBuilder(doc);
            builder.Writeln("Hello, Aspose.Words!");
            // Save the generated DOCX so the example can be rerun without errors.
            doc.Save(inputPath);
        }

        // Configure save options for MHTML with embedded resources.
        var saveOptions = new HtmlSaveOptions(SaveFormat.Mhtml)
        {
            CssStyleSheetType = CssStyleSheetType.External,
            ExportCidUrlsForMhtmlResources = true,
            ExportImagesAsBase64 = true,
            ExportFontsAsBase64 = true,
            PrettyFormat = true
        };

        // Save the document as MHTML.
        doc.Save(outputPath, saveOptions);
        Console.WriteLine($"MHTML file saved to: {outputPath}");
    }
}
