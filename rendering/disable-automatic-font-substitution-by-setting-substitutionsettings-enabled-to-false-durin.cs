using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fonts;

public class Program
{
    public static void Main()
    {
        // Define output paths.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);
        string pdfPath = Path.Combine(outputDir, "RenderedDocument.pdf");

        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Write some text using a font that is unlikely to exist on the system.
        builder.Font.Name = "MissingFont";
        builder.Writeln("This text uses a missing font. Font substitution is disabled, so the text will be rendered with the original metrics.");

        // Configure FontSettings to disable all font substitution rules.
        FontSettings fontSettings = new FontSettings();
        fontSettings.SubstitutionSettings.FontNameSubstitution.Enabled = false;
        fontSettings.SubstitutionSettings.FontConfigSubstitution.Enabled = false;
        fontSettings.SubstitutionSettings.TableSubstitution.Enabled = false;
        fontSettings.SubstitutionSettings.FontInfoSubstitution.Enabled = false;
        fontSettings.SubstitutionSettings.DefaultFontSubstitution.Enabled = false;

        // Assign the configured FontSettings to the document.
        doc.FontSettings = fontSettings;

        // Render the document to PDF.
        doc.Save(pdfPath, SaveFormat.Pdf);

        // Verify that the PDF file was created.
        if (!File.Exists(pdfPath))
            throw new InvalidOperationException("The PDF file was not created.");

        // Optionally, you could add further validation here (e.g., file size check).
    }
}
