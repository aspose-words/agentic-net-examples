using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Saving;

public class ExportFontsExample
{
    public static void Main()
    {
        // Prepare output folder.
        string outputFolder = Path.Combine(Directory.GetCurrentDirectory(), "ExportedFonts");
        Directory.CreateDirectory(outputFolder);

        // Path for the resulting HTML file.
        string htmlPath = Path.Combine(outputFolder, "document.html");

        // Create a simple document with some text.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Font.Name = "Arial";
        builder.Writeln("Sample text using Arial font.");

        // Configure HTML save options to export fonts as separate files.
        HtmlSaveOptions options = new HtmlSaveOptions
        {
            ExportFontResources = true,
            FontSavingCallback = new HandleFontSaving(outputFolder)
        };

        // Save the document to HTML.
        doc.Save(htmlPath, options);

        // Validate that the HTML file was created.
        if (!File.Exists(htmlPath))
            throw new InvalidOperationException("HTML file was not created.");

        // Validate that at least one font file was exported.
        string[] exportedFonts = Directory.GetFiles(outputFolder, "*.ttf");
        if (exportedFonts.Length == 0)
            throw new InvalidOperationException("No font files were exported.");

        // List exported font files.
        foreach (string fontFile in exportedFonts)
            Console.WriteLine($"Exported font: {fontFile}");
    }

    // Callback that controls how each font resource is saved.
    private class HandleFontSaving : IFontSavingCallback
    {
        private readonly string _outputFolder;

        public HandleFontSaving(string outputFolder)
        {
            _outputFolder = outputFolder;
        }

        void IFontSavingCallback.FontSaving(FontSavingArgs args)
        {
            // Use the original font file name for the exported file.
            string fileName = args.OriginalFileName.Split(Path.DirectorySeparatorChar).Last();
            args.FontFileName = fileName;

            // Save the font to a file in the output folder.
            string fullPath = Path.Combine(_outputFolder, fileName);
            args.FontStream = new FileStream(fullPath, FileMode.Create);
            args.KeepFontStreamOpen = false;
        }
    }
}
