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

        // Write some text using a font that is typically missing on Linux.
        builder.Font.Name = "Arial";
        builder.Writeln("This text is formatted with Arial, which should be substituted.");

        // Configure font substitution only on Linux/macOS platforms.
        PlatformID platform = Environment.OSVersion.Platform;
        bool isLinuxOrMac = platform == PlatformID.Unix || platform == PlatformID.MacOSX;

        if (isLinuxOrMac)
        {
            // Create FontSettings and add a substitution rule:
            // If "Arial" is not found, use "Liberation Sans" instead.
            FontSettings fontSettings = new FontSettings();
            fontSettings.SubstitutionSettings.TableSubstitution.AddSubstitutes("Arial", "Liberation Sans");

            // Assign the FontSettings to the document.
            doc.FontSettings = fontSettings;
        }

        // Define output path.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "FontSubstitutionOutput.pdf");

        // Save the document as PDF.
        doc.Save(outputPath, SaveFormat.Pdf);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("The output PDF was not created.");

        // Optionally, you could add further validation here (e.g., check file size).
    }
}
