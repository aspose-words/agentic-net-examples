using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fonts;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Use a font that supports ligatures (e.g., Times New Roman) and write text containing a ligature.
        builder.Font.Name = "Times New Roman";
        builder.Writeln("office"); // The "ff" sequence can be rendered as a ligature.

        // Configure FontSettings to point to the system fonts folder.
        FontSettings fontSettings = new FontSettings();
        string fontsFolder = Environment.GetFolderPath(Environment.SpecialFolder.Fonts);
        fontSettings.SetFontsFolder(fontsFolder, false);
        doc.FontSettings = fontSettings;

        // Render the document to a multipage TIFF image.
        string outputPath = "output.tiff";
        ImageSaveOptions saveOptions = new ImageSaveOptions(SaveFormat.Tiff);
        doc.Save(outputPath, saveOptions);

        // Verify that the TIFF file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("Failed to create the TIFF file.");

        // Output the result path.
        Console.WriteLine($"TIFF file saved to: {Path.GetFullPath(outputPath)}");
    }
}
