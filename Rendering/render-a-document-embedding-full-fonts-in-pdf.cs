using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Words.Fonts;

class EmbedFullFontsPdfExample
{
    static void Main()
    {
        // Define the output folder and ensure it exists.
        string artifactsDir = Path.Combine(Environment.CurrentDirectory, "Output");
        Directory.CreateDirectory(artifactsDir);

        // Create a new empty document.
        Document doc = new Document();

        // Use DocumentBuilder to add some text with a custom font.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Font.Name = "Arial";
        builder.Writeln("Hello world!");
        builder.Font.Name = "Times New Roman";
        builder.Writeln("The quick brown fox jumps over the lazy dog.");

        // Optionally, configure additional font sources if needed.
        // This ensures that the fonts used above are available during rendering.
        FontSourceBase[] originalSources = FontSettings.DefaultInstance.GetFontsSources();
        FolderFontSource folderSource = new FolderFontSource(@"C:\Windows\Fonts", true);
        FontSettings.DefaultInstance.SetFontsSources(new[] { originalSources[0], folderSource });

        // Create PdfSaveOptions and enable full font embedding.
        PdfSaveOptions options = new PdfSaveOptions
        {
            EmbedFullFonts = true // Embed the complete font files without subsetting.
        };

        // Save the document as PDF with the specified options.
        string outputPath = Path.Combine(artifactsDir, "DocumentWithFullFonts.pdf");
        doc.Save(outputPath, options);

        // Restore the original font sources (optional cleanup).
        FontSettings.DefaultInstance.SetFontsSources(originalSources);

        Console.WriteLine($"PDF saved to: {outputPath}");
    }
}
