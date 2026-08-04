using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fonts;

public class Program
{
    public static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();

        // Initialize FontSettings for the document.
        FontSettings fontSettings = new FontSettings();
        doc.FontSettings = fontSettings;

        // Add a custom substitution: replace missing "Times New Roman" with "Calibri".
        // This uses the table substitution rule.
        fontSettings.SubstitutionSettings.TableSubstitution.AddSubstitutes(
            "Times New Roman", "Calibri");

        // Build document content using a font that may be missing.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Font.Name = "Times New Roman";
        builder.Writeln("This line uses Times New Roman, which will be substituted with Calibri if unavailable.");

        // Define output path.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "FontSubstitution.pdf");

        // Save the document as PDF.
        doc.Save(outputPath);
    }
}
