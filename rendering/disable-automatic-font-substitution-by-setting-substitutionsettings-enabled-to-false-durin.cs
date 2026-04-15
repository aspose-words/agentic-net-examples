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
        string outputPath = Path.Combine(outputDir, "NoFontSubstitution.pdf");

        // Create a simple document with text using a font that likely does not exist.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Font.Name = "MissingFont";
        builder.Writeln("This text is formatted with a missing font.");

        // Configure font settings to disable all automatic substitution rules.
        FontSettings fontSettings = new FontSettings();

        // Disable default font substitution rule.
        fontSettings.SubstitutionSettings.DefaultFontSubstitution.Enabled = false;

        // Disable font name substitution rule.
        fontSettings.SubstitutionSettings.FontNameSubstitution.Enabled = false;

        // Disable font info substitution rule.
        fontSettings.SubstitutionSettings.FontInfoSubstitution.Enabled = false;

        // Disable font config substitution rule (if available on the platform).
        fontSettings.SubstitutionSettings.FontConfigSubstitution.Enabled = false;

        // Assign the configured settings to the document.
        doc.FontSettings = fontSettings;

        // Render the document to PDF.
        doc.Save(outputPath, SaveFormat.Pdf);

        // Verify that the PDF file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("The PDF file was not created.");

        // Optionally, output the result path (no interactive prompts required).
        Console.WriteLine($"PDF saved to: {outputPath}");
    }
}
