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
        string outputPath = Path.Combine(outputDir, "CustomFontSubstitution.pdf");

        // Create a new blank document.
        Document doc = new Document();

        // Configure font settings.
        FontSettings fontSettings = new FontSettings();
        doc.FontSettings = fontSettings;

        // Add a custom substitution: replace missing "Times New Roman" with "Calibri".
        TableSubstitutionRule tableRule = fontSettings.SubstitutionSettings.TableSubstitution;
        tableRule.AddSubstitutes("Times New Roman", "Calibri");

        // Validate that the substitution was added.
        var substitutes = tableRule.GetSubstitutes("Times New Roman");
        if (substitutes == null || !substitutes.Contains("Calibri"))
            throw new InvalidOperationException("Failed to add Calibri as a substitute for Times New Roman.");

        // Write some text using the font that we want to substitute.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Font.Name = "Times New Roman";
        builder.Writeln("This line is formatted with Times New Roman, which will be rendered using Calibri.");

        // Save the document to PDF.
        doc.Save(outputPath);

        // Ensure the file was created.
        if (!File.Exists(outputPath))
            throw new FileNotFoundException("The output PDF was not created.", outputPath);
    }
}
