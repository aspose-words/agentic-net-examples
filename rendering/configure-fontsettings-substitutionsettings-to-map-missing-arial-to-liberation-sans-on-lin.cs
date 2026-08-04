using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fonts;

public class Program
{
    public static void Main()
    {
        // Prepare output folder.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Write some text using a font that may be missing on Linux.
        builder.Font.Name = "Arial";
        builder.Writeln("This text is formatted with Arial. On Linux it should be substituted with Liberation Sans.");

        // Configure font substitution: map missing Arial to Liberation Sans on Linux platforms.
        FontSettings fontSettings = new FontSettings();

        PlatformID platform = Environment.OSVersion.Platform;
        bool isLinux = platform == PlatformID.Unix || platform == PlatformID.MacOSX; // Treat macOS similarly if needed.

        if (isLinux)
        {
            // Add a substitute font for Arial.
            fontSettings.SubstitutionSettings.TableSubstitution.AddSubstitutes("Arial", "Liberation Sans");
        }

        // Assign the configured FontSettings to the document.
        doc.FontSettings = fontSettings;

        // Save the document to PDF to trigger rendering and substitution.
        string outputPath = Path.Combine(outputDir, "Result.pdf");
        doc.Save(outputPath);

        // Verify that the output file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("Failed to create the output PDF.");

        // Optionally, indicate success (no interactive prompts required).
        Console.WriteLine("Document saved to: " + outputPath);
    }
}
