using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Prepare folders.
        string baseDir = Directory.GetCurrentDirectory();
        string artifactsDir = Path.Combine(baseDir, "Artifacts");
        Directory.CreateDirectory(artifactsDir);

        // 1. Create a sample DOCX document.
        string inputDocx = Path.Combine(artifactsDir, "Sample.docx");
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);
        builder.Writeln("Sample text with default font.");
        builder.Font.Name = "Arial";
        builder.Writeln("This line uses Arial.");
        builder.Font.Name = "Times New Roman";
        builder.Writeln("This line uses Times New Roman.");
        sourceDoc.Save(inputDocx, SaveFormat.Docx);

        // 2. Load the document we just created.
        Document doc = new Document(inputDocx);

        // 3. Configure HTML save options to export fonts as separate files.
        HtmlSaveOptions saveOptions = new HtmlSaveOptions
        {
            ExportFontResources = true,
            FontSavingCallback = new HandleFontSaving(artifactsDir)
        };

        // 4. Save the document as HTML. Fonts will be written to the same folder.
        string htmlPath = Path.Combine(artifactsDir, "Sample.html");
        doc.Save(htmlPath, saveOptions);

        // 5. Validate output.
        if (!File.Exists(htmlPath))
            throw new InvalidOperationException("HTML file was not created.");

        string[] fontFiles = Directory.GetFiles(artifactsDir, "*.ttf");
        if (fontFiles.Length == 0)
            throw new InvalidOperationException("No font files were exported.");

        // 6. List the generated files.
        Console.WriteLine("HTML file created: " + htmlPath);
        Console.WriteLine("Exported font files:");
        foreach (string fontFile in fontFiles)
            Console.WriteLine("  " + Path.GetFileName(fontFile));
    }
}

// Callback that tells Aspose.Words how to save each exported font.
public class HandleFontSaving : IFontSavingCallback
{
    private readonly string _outputFolder;

    public HandleFontSaving(string outputFolder)
    {
        _outputFolder = outputFolder;
    }

    void IFontSavingCallback.FontSaving(FontSavingArgs args)
    {
        // Use only the file name part of the original font file.
        string fontFileName = args.OriginalFileName.Split(Path.DirectorySeparatorChar).Last();

        // Set the file name Aspose.Words will use.
        args.FontFileName = fontFileName;

        // Provide a stream that writes the font into our output folder.
        string fontPath = Path.Combine(_outputFolder, fontFileName);
        args.FontStream = new FileStream(fontPath, FileMode.Create);

        // Ensure Aspose.Words closes the stream after writing.
        args.KeepFontStreamOpen = false;
    }
}
