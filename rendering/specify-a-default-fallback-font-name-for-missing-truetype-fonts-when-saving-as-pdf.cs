using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fonts;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Define output folder and file name.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);
        string pdfPath = Path.Combine(outputDir, "DocumentWithFallback.pdf");

        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Use a font that does not exist on the system.
        builder.Font.Name = "NonExistentFont";
        builder.Writeln("This text uses a missing font and will be substituted.");

        // Configure font substitution: set a default fallback font.
        FontSettings fontSettings = new FontSettings();
        DefaultFontSubstitutionRule defaultSubstitution = fontSettings.SubstitutionSettings.DefaultFontSubstitution;
        defaultSubstitution.Enabled = true;               // Ensure the rule is active.
        defaultSubstitution.DefaultFontName = "Courier New"; // Fallback font to use.

        // Assign the FontSettings to the document.
        doc.FontSettings = fontSettings;

        // Save the document as PDF.
        doc.Save(pdfPath, SaveFormat.Pdf);

        // Verify that the PDF file was created.
        if (!File.Exists(pdfPath))
            throw new FileNotFoundException("PDF file was not created.", pdfPath);
    }
}
