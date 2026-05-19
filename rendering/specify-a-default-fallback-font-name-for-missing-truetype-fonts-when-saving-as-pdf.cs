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

        // Use a font that does not exist on the system.
        builder.Font.Name = "NonExistentFont";
        builder.Writeln("This text uses a missing font and will be substituted.");

        // Configure default font substitution rule.
        FontSettings fontSettings = new FontSettings();
        DefaultFontSubstitutionRule defaultSubstitution = fontSettings.SubstitutionSettings.DefaultFontSubstitution;
        defaultSubstitution.Enabled = true;
        defaultSubstitution.DefaultFontName = "Times New Roman"; // fallback font

        // Assign the font settings to the document.
        doc.FontSettings = fontSettings;

        // Save the document as PDF.
        string pdfPath = Path.Combine(artifactsDir, "FallbackFontExample.pdf");
        doc.Save(pdfPath, SaveFormat.Pdf);

        // Verify that the PDF file was created.
        if (!File.Exists(pdfPath))
            throw new InvalidOperationException("PDF file was not created.");

        // Simple confirmation (no interactive output required).
        Console.WriteLine("PDF saved to: " + pdfPath);
    }
}
