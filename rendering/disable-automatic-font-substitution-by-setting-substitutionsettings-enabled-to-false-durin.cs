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
        string pdfPath = Path.Combine(outputDir, "NoFontSubstitution.pdf");

        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Write text using a font that is unlikely to exist on the system.
        builder.Font.Name = "MissingFont";
        builder.Writeln("This text is formatted with a missing font.");

        // Configure font settings and disable all substitution rules.
        FontSettings fontSettings = new FontSettings();
        fontSettings.SubstitutionSettings.FontNameSubstitution.Enabled = false;
        fontSettings.SubstitutionSettings.FontConfigSubstitution.Enabled = false;
        fontSettings.SubstitutionSettings.TableSubstitution.Enabled = false;
        fontSettings.SubstitutionSettings.FontInfoSubstitution.Enabled = false;
        fontSettings.SubstitutionSettings.DefaultFontSubstitution.Enabled = false;

        // Apply the font settings to the document.
        doc.FontSettings = fontSettings;

        // Render the document to PDF.
        doc.Save(pdfPath);

        // Verify that the PDF was created.
        if (!File.Exists(pdfPath))
            throw new InvalidOperationException("The PDF file was not created.");

        // Optionally, output the path (no interactive prompts required).
        Console.WriteLine($"PDF saved to: {pdfPath}");
    }
}
