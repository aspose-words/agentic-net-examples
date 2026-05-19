using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Output folder for HTML and exported fonts
        string outputFolder = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputFolder);

        // Create a sample document that uses a couple of different fonts
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Font.Name = "Arial";
        builder.Writeln("This paragraph uses Arial font.");
        builder.Font.Name = "Times New Roman";
        builder.Writeln("This paragraph uses Times New Roman font.");

        // Configure HTML save options to export fonts as separate files
        HtmlSaveOptions saveOptions = new HtmlSaveOptions
        {
            ExportFontResources = true,
            FontSavingCallback = new HandleFontSaving(outputFolder)
        };

        // Save the document as HTML
        string htmlPath = Path.Combine(outputFolder, "Document.html");
        doc.Save(htmlPath, saveOptions);

        // Verify that at least one .ttf file was exported
        string[] exportedFonts = Directory.GetFiles(outputFolder, "*.ttf");
        if (exportedFonts.Length == 0)
            throw new InvalidOperationException("No font files were exported.");

        // List the exported font files
        foreach (string fontFile in exportedFonts)
            Console.WriteLine(fontFile);
    }
}

// Callback that controls how each font resource is saved
public class HandleFontSaving : IFontSavingCallback
{
    private readonly string _folder;

    public HandleFontSaving(string folder)
    {
        _folder = folder;
    }

    void IFontSavingCallback.FontSaving(FontSavingArgs args)
    {
        // Use the original font file name (without path) for the exported file
        string fileName = Path.GetFileName(args.OriginalFileName);
        args.FontFileName = fileName;

        // Save the font to a file in the specified folder
        string fullPath = Path.Combine(_folder, fileName);
        args.FontStream = new FileStream(fullPath, FileMode.Create);
        args.KeepFontStreamOpen = false;
    }
}
