using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fonts;

public class Program
{
    public static void Main()
    {
        // Create a simple document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Sample text using custom FontSettings.");

        // Create FontSettings and point it to a network folder containing TrueType fonts.
        FontSettings fontSettings = new FontSettings();
        // Example UNC path to a network share. Replace with an actual network path if needed.
        string networkFontsFolder = @"\\MyNetworkShare\Fonts";
        // Set the folder; false indicates not to search subfolders.
        fontSettings.SetFontsFolder(networkFontsFolder, false);

        // Assign the FontSettings to the document.
        doc.FontSettings = fontSettings;

        // Save the document to verify that FontSettings were applied without errors.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "Output.docx");
        doc.Save(outputPath, SaveFormat.Docx);

        // Ensure the output file was created.
        if (!File.Exists(outputPath))
        {
            throw new InvalidOperationException("The document was not saved successfully.");
        }
    }
}
