using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fonts;

public class Program
{
    public static void Main()
    {
        // Prepare output folder.
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        Directory.CreateDirectory(artifactsDir);

        // Create a simple document with text using a font that is unlikely to be present.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Font.Name = "MissingFont";
        builder.Writeln("This text uses a missing font. Font substitution is disabled.");

        // Configure font settings to disable all substitution rules.
        FontSettings fontSettings = new FontSettings();
        fontSettings.SubstitutionSettings.FontNameSubstitution.Enabled = false;
        fontSettings.SubstitutionSettings.FontConfigSubstitution.Enabled = false;
        fontSettings.SubstitutionSettings.TableSubstitution.Enabled = false;
        fontSettings.SubstitutionSettings.FontInfoSubstitution.Enabled = false;
        fontSettings.SubstitutionSettings.DefaultFontSubstitution.Enabled = false;
        doc.FontSettings = fontSettings;

        // Render the document to PDF.
        string pdfPath = Path.Combine(artifactsDir, "NoSubstitution.pdf");
        doc.Save(pdfPath);

        // Verify that the PDF was created.
        if (!File.Exists(pdfPath))
            throw new InvalidOperationException("Failed to create the PDF file.");
    }
}
