using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Words.Fonts;

public class Program
{
    public static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Use a font that supports ligatures (e.g., Calibri).
        builder.Font.Name = "Calibri";

        // Add a line containing ligature characters.
        builder.Writeln("Office: fi fl ffi ffl");

        // Configure FontSettings to point to the system fonts folder.
        FontSettings fontSettings = new FontSettings();
        string fontsFolder = Environment.GetFolderPath(Environment.SpecialFolder.Fonts);
        if (Directory.Exists(fontsFolder))
        {
            fontSettings.SetFontsFolder(fontsFolder, false);
        }
        doc.FontSettings = fontSettings;

        // Render the document to a TIFF image.
        string outputPath = "output.tiff";
        ImageSaveOptions saveOptions = new ImageSaveOptions(SaveFormat.Tiff);
        doc.Save(outputPath, saveOptions);

        // Verify that the TIFF file was created.
        if (!File.Exists(outputPath))
        {
            throw new InvalidOperationException("TIFF file was not created.");
        }

        // Output the file size for confirmation.
        Console.WriteLine($"TIFF saved successfully. Size: {new FileInfo(outputPath).Length} bytes");
    }
}
