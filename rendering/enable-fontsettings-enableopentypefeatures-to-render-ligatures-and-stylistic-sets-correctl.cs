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

        // Use a font that supports ligatures (e.g., Times New Roman).
        builder.Font.Name = "Times New Roman";

        // Add text that contains ligatures such as "ff" and "fl".
        builder.Writeln("Office");   // contains "ff"
        builder.Writeln("flavor");   // contains "fl"

        // Configure FontSettings to point to the system fonts folder.
        // This ensures the required font can be found during rendering.
        FontSettings fontSettings = new FontSettings();
        string systemFontsPath = Environment.GetFolderPath(Environment.SpecialFolder.Fonts);
        if (Directory.Exists(systemFontsPath))
        {
            fontSettings.SetFontsFolder(systemFontsPath, false);
        }
        doc.FontSettings = fontSettings;

        // Define the output TIFF file path.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "RenderedDocument.tiff");

        // Set up image save options for TIFF format.
        ImageSaveOptions saveOptions = new ImageSaveOptions(SaveFormat.Tiff);
        // Render all pages of the document.
        saveOptions.PageSet = new PageSet(0, doc.PageCount - 1);

        // Save the document as a multipage TIFF.
        doc.Save(outputPath, saveOptions);

        // Verify that the TIFF file was created.
        if (!File.Exists(outputPath))
        {
            throw new InvalidOperationException("TIFF rendering failed: output file was not created.");
        }

        // Report success.
        Console.WriteLine("Document rendered to TIFF successfully: " + outputPath);
    }
}
