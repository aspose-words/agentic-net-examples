using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fonts;

public class Program
{
    public static void Main()
    {
        // Prepare output directory.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);
        string outputPath = Path.Combine(outputDir, "MappedArial.pdf");

        // Create a simple document with text formatted in Arial.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Font.Name = "Arial";
        builder.Writeln("This text is formatted with Arial. If Arial is missing, it should be substituted with Liberation Sans.");

        // Configure font substitution only on Linux/macOS platforms.
        PlatformID platform = Environment.OSVersion.Platform;
        bool isLinuxOrMac = platform == PlatformID.Unix || platform == PlatformID.MacOSX;
        if (isLinuxOrMac)
        {
            FontSettings fontSettings = new FontSettings();
            // Map missing "Arial" to "Liberation Sans".
            fontSettings.SubstitutionSettings.TableSubstitution.SetSubstitutes("Arial", "Liberation Sans");
            doc.FontSettings = fontSettings;
        }

        // Save the document as PDF.
        doc.Save(outputPath, SaveFormat.Pdf);

        // Verify that the output file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException($"Failed to create output file: {outputPath}");
    }
}
