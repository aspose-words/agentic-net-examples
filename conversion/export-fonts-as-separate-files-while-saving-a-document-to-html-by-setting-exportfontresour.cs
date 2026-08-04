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

        // Create a simple document with some text.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Font.Name = "Arial";
        builder.Writeln("Sample text for font export demonstration.");

        // Configure HTML save options to export fonts as separate files.
        HtmlSaveOptions options = new HtmlSaveOptions
        {
            ExportFontResources = true,
            FontsFolder = outputDir,               // Folder where fonts will be written.
            FontSavingCallback = new HandleFontSaving()
        };

        // Save the document to HTML using the configured options.
        string htmlPath = Path.Combine(outputDir, "output.html");
        doc.Save(htmlPath, options);

        // Verify that at least one font file was exported.
        string[] fontFiles = Directory.GetFiles(outputDir, "*.ttf");
        if (fontFiles.Length == 0)
            throw new InvalidOperationException("No font files were exported.");

        // List exported font files.
        foreach (string fontFile in fontFiles)
            Console.WriteLine($"Exported font: {Path.GetFileName(fontFile)}");
    }

    // Callback that controls how each font resource is saved.
    private class HandleFontSaving : IFontSavingCallback
    {
        void IFontSavingCallback.FontSaving(FontSavingArgs args)
        {
            // Use the original font file name for the exported file.
            string fontFileName = args.OriginalFileName.Split(Path.DirectorySeparatorChar).Last();
            args.FontFileName = fontFileName;

            // Save the font to a file in the same output folder.
            string fontPath = Path.Combine(Directory.GetCurrentDirectory(), "Output", fontFileName);
            args.FontStream = new FileStream(fontPath, FileMode.Create);
            args.KeepFontStreamOpen = false;
        }
    }
}
