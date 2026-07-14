using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fonts;

public class FontSubstitutionExample
{
    public static void Main()
    {
        // Prepare output directory.
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        Directory.CreateDirectory(artifactsDir);

        // Create a new blank document.
        Document doc = new Document();

        // Initialize FontSettings for the document.
        FontSettings fontSettings = new FontSettings();
        doc.FontSettings = fontSettings;

        // Define a custom substitution: replace missing "Times New Roman" with "Calibri".
        // This rule will be applied only when "Times New Roman" cannot be found.
        fontSettings.SubstitutionSettings.TableSubstitution.SetSubstitutes("Times New Roman", "Calibri");

        // Build document content using a font that is intentionally missing.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Font.Name = "Times New Roman"; // This font will be substituted.
        builder.Writeln("This line is written with Times New Roman, which will be rendered as Calibri.");

        // Save the document to PDF to observe the substitution.
        string outputPath = Path.Combine(artifactsDir, "FontSubstitution.pdf");
        doc.Save(outputPath);
    }
}
