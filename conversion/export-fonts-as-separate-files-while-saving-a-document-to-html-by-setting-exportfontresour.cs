using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Prepare output folder.
        string outputFolder = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputFolder);

        // Create a sample document with different fonts.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.Font.Name = "Arial";
        builder.Writeln("This text uses Arial font.");

        builder.Font.Name = "Times New Roman";
        builder.Writeln("This text uses Times New Roman font.");

        // Configure HTML save options to export fonts as separate files.
        HtmlSaveOptions saveOptions = new HtmlSaveOptions
        {
            ExportFontResources = true,
            FontSavingCallback = new HandleFontSaving(outputFolder)
        };

        // Save the document as HTML.
        string htmlPath = Path.Combine(outputFolder, "Document.html");
        doc.Save(htmlPath, saveOptions);

        // Verify that at least one font file was exported.
        string[] fontFiles = Directory.GetFiles(outputFolder, "*.ttf");
        if (fontFiles.Length == 0)
        {
            throw new InvalidOperationException("No font files were exported.");
        }

        // List exported font files.
        Console.WriteLine("Exported font files:");
        foreach (string fontFile in fontFiles)
        {
            Console.WriteLine(fontFile);
        }

        // Optionally, list the generated HTML file.
        Console.WriteLine($"HTML document saved to: {htmlPath}");
    }
}

// Implements custom logic for saving exported fonts.
public class HandleFontSaving : IFontSavingCallback
{
    private readonly string _outputFolder;

    public HandleFontSaving(string outputFolder)
    {
        _outputFolder = outputFolder ?? throw new ArgumentNullException(nameof(outputFolder));
    }

    void IFontSavingCallback.FontSaving(FontSavingArgs args)
    {
        // Use only the file name part of the original font file.
        string fontFileName = Path.GetFileName(args.OriginalFileName);
        args.FontFileName = fontFileName;

        // Save the font to the output folder.
        string fontPath = Path.Combine(_outputFolder, fontFileName);
        args.FontStream = new FileStream(fontPath, FileMode.Create);
        args.KeepFontStreamOpen = false;
    }
}
