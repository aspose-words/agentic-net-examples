using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fonts;

public class Program
{
    public static void Main()
    {
        // Define folders for the example.
        string baseDir = Directory.GetCurrentDirectory();
        string outputDir = Path.Combine(baseDir, "Output");
        string fontDir = Path.Combine(baseDir, "NetworkFonts");

        // Ensure the directories exist.
        Directory.CreateDirectory(outputDir);
        Directory.CreateDirectory(fontDir);

        // Simulate a network folder using a UNC path.
        // In a real scenario this would be something like "\\\\Server\\Share\\Fonts".
        // Here we use a local UNC path that points to the folder we just created.
        string networkFontPath = @"\\127.0.0.1\" + fontDir.TrimStart(Path.GetPathRoot(fontDir).ToCharArray());

        // Create a simple document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Font.Name = "CustomFont"; // Use a font that may not be installed.
        builder.Writeln("This text is rendered using a custom font source.");

        // Configure custom FontSettings to point to the network folder.
        FontSettings fontSettings = new FontSettings();
        // The second argument indicates whether to scan subfolders.
        fontSettings.SetFontsFolder(networkFontPath, true);
        doc.FontSettings = fontSettings;

        // Save the document.
        string outputPath = Path.Combine(outputDir, "CustomFontSettings.pdf");
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("The document was not saved successfully.");

        // Indicate completion.
        Console.WriteLine("Document saved to: " + outputPath);
    }
}
