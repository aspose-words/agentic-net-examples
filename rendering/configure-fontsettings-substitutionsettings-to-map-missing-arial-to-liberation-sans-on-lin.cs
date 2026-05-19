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

        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Write some text using the Arial font (which may be missing on Linux).
        builder.Font.Name = "Arial";
        builder.Writeln("This text is formatted with Arial. If Arial is unavailable, it should be substituted with Liberation Sans.");

        // Configure font substitution.
        FontSettings fontSettings = new FontSettings();

        // Apply substitution only on non‑Windows platforms.
        PlatformID platform = Environment.OSVersion.Platform;
        bool isLinuxOrMac = platform == PlatformID.Unix || platform == PlatformID.MacOSX;
        if (isLinuxOrMac)
        {
            // Map missing Arial to Liberation Sans.
            fontSettings.SubstitutionSettings.TableSubstitution.SetSubstitutes("Arial", "Liberation Sans");
        }

        // Assign the configured FontSettings to the document.
        doc.FontSettings = fontSettings;

        // Save the document as PDF.
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new Exception("Failed to create the output PDF file.");

        // Optionally, you could add further processing here.
    }
}
