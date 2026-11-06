using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fonts;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Configure FontSettings to use the system fonts folder.
        string fontsFolder = Environment.GetFolderPath(Environment.SpecialFolder.Fonts);
        FontSettings fontSettings = new FontSettings();
        fontSettings.SetFontsFolder(fontsFolder, false);
        doc.FontSettings = fontSettings;

        // Add text that contains ligatures (e.g., "ff", "fi") and Arabic text for contextual forms.
        builder.Font.Name = "Times New Roman";
        builder.Writeln("Office");      // Contains "ff" ligature.
        builder.Writeln("affinity");    // Contains "fi" ligature.

        builder.Font.Name = "Arial";
        builder.Writeln("العربية");      // Arabic text with contextual shaping.

        // Render the document to PDF. OpenType features (ligatures, contextual forms) are applied by default.
        string outputPath = "Output.pdf";
        doc.Save(outputPath, SaveFormat.Pdf);

        // Verify that the PDF file was created.
        if (!File.Exists(outputPath))
        {
            throw new InvalidOperationException($"Failed to create PDF file at '{outputPath}'.");
        }

        // Optionally, you could inspect the PDF bytes for embedded font markers,
        // but for this example we only confirm the file exists.
        Console.WriteLine($"PDF rendering completed successfully. File saved to: {Path.GetFullPath(outputPath)}");
    }
}
