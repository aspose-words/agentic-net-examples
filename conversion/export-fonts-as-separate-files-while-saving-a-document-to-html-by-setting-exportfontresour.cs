using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Prepare output folder.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // Create a simple document with a specific font.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Font.Name = "Arial";
        builder.Writeln("This text uses the Arial font and will trigger font export.");

        // Configure HTML save options to export fonts as separate files.
        HtmlSaveOptions saveOptions = new HtmlSaveOptions
        {
            ExportFontResources = true,
            FontsFolder = outputDir,               // Folder where font files will be written.
            FontSavingCallback = new HandleFontSaving()
        };

        // Save the document as HTML.
        string htmlPath = Path.Combine(outputDir, "sample.html");
        doc.Save(htmlPath, saveOptions);

        // Validate that the HTML file was created.
        if (!File.Exists(htmlPath))
            throw new InvalidOperationException("HTML output file was not created.");

        // Validate that at least one font file (e.g., .ttf) was exported.
        string[] exportedFonts = Directory.GetFiles(outputDir, "*.ttf");
        if (exportedFonts.Length == 0)
            throw new InvalidOperationException("No font files were exported.");

        // Output the locations of the generated files.
        Console.WriteLine($"HTML file saved to: {htmlPath}");
        foreach (string fontFile in exportedFonts)
        {
            Console.WriteLine($"Exported font: {fontFile}");
        }
    }

    // Callback that controls how each font resource is saved.
    private class HandleFontSaving : IFontSavingCallback
    {
        void IFontSavingCallback.FontSaving(FontSavingArgs args)
        {
            // Use the original font file name (without path) for the exported file.
            args.FontFileName = args.OriginalFileName.Split(Path.DirectorySeparatorChar).Last();

            // Optionally, you could provide a custom stream:
            // args.FontStream = new FileStream(Path.Combine(Directory.GetCurrentDirectory(), "Output", args.FontFileName), FileMode.Create);
            // args.KeepFontStreamOpen = false;
        }
    }
}
