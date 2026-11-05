using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fonts;

public class Program
{
    public static void Main()
    {
        // Create a new document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Set a primary font that does not contain certain Unicode characters.
        builder.Font.Name = "Times New Roman";
        builder.Writeln("Hello, world! 你好, مرحبا");

        // Configure font substitution to use Arial Unicode MS as fallback.
        FontSettings fontSettings = new FontSettings();
        fontSettings.SubstitutionSettings.DefaultFontSubstitution.DefaultFontName = "Arial Unicode MS";
        doc.FontSettings = fontSettings;

        // Render the document to PDF.
        string outputPath = "Output.pdf";
        doc.Save(outputPath, SaveFormat.Pdf);

        // Verify that the PDF was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("Failed to create the PDF output.");

        // Optionally, indicate success (no interactive input required).
        Console.WriteLine("PDF generated successfully with font fallback.");
    }
}
