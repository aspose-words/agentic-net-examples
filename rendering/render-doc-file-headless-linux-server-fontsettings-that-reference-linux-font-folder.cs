using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fonts;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Create a simple document in memory.
        var doc = new Document();
        var builder = new DocumentBuilder(doc);
        builder.Writeln("Hello, Aspose.Words!");

        // Linux font folder that contains TrueType fonts.
        // If the folder does not exist (e.g., on Windows), the setting is simply ignored.
        string linuxFontsFolder = "/usr/share/fonts";

        if (Directory.Exists(linuxFontsFolder))
        {
            var fontSettings = new FontSettings();
            // Search recursively for fonts in the specified folder.
            fontSettings.SetFontsFolder(linuxFontsFolder, true);
            doc.FontSettings = fontSettings;
        }

        // Determine output path relative to the current working directory.
        string outputPdfPath = Path.Combine(Directory.GetCurrentDirectory(), "output.pdf");

        // Save the document as PDF.
        doc.Save(outputPdfPath, SaveFormat.Pdf);

        Console.WriteLine($"PDF saved to: {outputPdfPath}");
    }
}
