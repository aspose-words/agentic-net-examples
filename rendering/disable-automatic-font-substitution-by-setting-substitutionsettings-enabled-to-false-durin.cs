using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fonts;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Prepare output directory.
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        Directory.CreateDirectory(artifactsDir);

        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Write some text using a font that is unlikely to be present.
        builder.Font.Name = "MissingFont";
        builder.Writeln("This text uses a missing font. No substitution should occur.");

        // Configure font settings to disable all substitution rules.
        FontSettings fontSettings = new FontSettings();
        fontSettings.SubstitutionSettings.FontNameSubstitution.Enabled = false;
        fontSettings.SubstitutionSettings.FontConfigSubstitution.Enabled = false;
        fontSettings.SubstitutionSettings.TableSubstitution.Enabled = false;
        fontSettings.SubstitutionSettings.FontInfoSubstitution.Enabled = false;
        fontSettings.SubstitutionSettings.DefaultFontSubstitution.Enabled = false;

        // Assign the configured settings to the document.
        doc.FontSettings = fontSettings;

        // Render the document to PDF.
        string outputPath = Path.Combine(artifactsDir, "NoSubstitution.pdf");
        doc.Save(outputPath, SaveFormat.Pdf);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("The PDF file was not created.");

        // Optionally, you could inspect the PDF for substitution markers,
        // but for this example we only ensure the file exists.
    }
}
