using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fonts;

public class Program
{
    public static void Main()
    {
        // Define a folder for output files.
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        Directory.CreateDirectory(artifactsDir);
        string outputPath = Path.Combine(artifactsDir, "CustomFontSubstitution.pdf");

        // Create a new empty document.
        Document doc = new Document();

        // Create a DocumentBuilder to add content.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Configure font settings with a table substitution rule:
        // If "Times New Roman" is missing, substitute it with "Calibri".
        FontSettings fontSettings = new FontSettings();
        fontSettings.SubstitutionSettings.TableSubstitution.AddSubstitutes("Times New Roman", "Calibri");
        doc.FontSettings = fontSettings;

        // Write a line using the font that may be missing.
        builder.Font.Name = "Times New Roman";

        // Validate that the font name was set correctly.
        if (builder.Font.Name != "Times New Roman")
        {
            throw new InvalidOperationException("Failed to set the font name on the builder.");
        }

        builder.Writeln("This text uses Times New Roman, which will be substituted with Calibri if unavailable.");

        // Save the document to PDF.
        doc.Save(outputPath);

        // Verify that the output file was created.
        if (!File.Exists(outputPath))
        {
            throw new FileNotFoundException("The output PDF was not created.", outputPath);
        }
    }
}
