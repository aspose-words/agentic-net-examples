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
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);
        string pdfPath = Path.Combine(outputDir, "Rendered.pdf");

        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Write text using a font that is unlikely to be present on the system.
        builder.Font.Name = "MissingFont";
        builder.Writeln("This text uses a missing font. Font substitution is disabled.");

        // Configure font settings to disable all substitution rules.
        FontSettings fontSettings = new FontSettings();

        // Disable each substitution rule.
        fontSettings.SubstitutionSettings.FontNameSubstitution.Enabled = false;
        fontSettings.SubstitutionSettings.FontInfoSubstitution.Enabled = false;
        fontSettings.SubstitutionSettings.TableSubstitution.Enabled = false;
        fontSettings.SubstitutionSettings.FontConfigSubstitution.Enabled = false;
        fontSettings.SubstitutionSettings.DefaultFontSubstitution.Enabled = false;

        // Assign the configured settings to the document.
        doc.FontSettings = fontSettings;

        // Render the document to PDF.
        PdfSaveOptions saveOptions = new PdfSaveOptions();
        doc.Save(pdfPath, saveOptions);

        // Verify that the PDF file was created.
        if (!File.Exists(pdfPath))
            throw new InvalidOperationException("The PDF file was not created.");

        // Optionally, you could add further validation here (e.g., file size check).
    }
}
