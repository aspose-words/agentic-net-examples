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

        // Write some text using a font that may be missing on Linux.
        builder.Font.Name = "Arial";
        builder.Writeln("This text is formatted with Arial. On Linux it will be substituted.");

        // Prepare font settings.
        FontSettings fontSettings = new FontSettings();

        // Apply substitution only on non‑Windows platforms (Linux/macOS).
        PlatformID platform = Environment.OSVersion.Platform;
        bool isLinuxOrMac = platform == PlatformID.Unix || platform == PlatformID.MacOSX;
        if (isLinuxOrMac)
        {
            // Map missing Arial to Liberation Sans.
            fontSettings.SubstitutionSettings.TableSubstitution.AddSubstitutes("Arial", "Liberation Sans");
        }

        // Assign the configured settings to the document.
        doc.FontSettings = fontSettings;

        // Define output path.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);
        string outputPath = Path.Combine(outputDir, "FontSubstitutionExample.pdf");

        // Save the document (PDF rendering).
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("Failed to create the output PDF file.");

        // Optionally, you could inspect the PDF for font substitution markers,
        // but for this example we only ensure the file exists.
    }
}
