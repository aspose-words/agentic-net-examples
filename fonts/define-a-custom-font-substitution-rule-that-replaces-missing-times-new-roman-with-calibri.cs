using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fonts;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Initialize FontSettings for the document.
        FontSettings fontSettings = new FontSettings();
        doc.FontSettings = fontSettings;

        // Define a custom substitution: if "Times New Roman" is missing, use "Calibri".
        // This uses the TableSubstitutionRule which allows specifying substitutes for a particular font name.
        doc.FontSettings.SubstitutionSettings.TableSubstitution.AddSubstitutes("Times New Roman", "Calibri");

        // Build the document content using a font that may be missing.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Font.Name = "Times New Roman";
        builder.Writeln("This line is formatted with Times New Roman, which will be substituted with Calibri if unavailable.");

        // Define output path.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);
        string outputPath = Path.Combine(outputDir, "CustomFontSubstitution.pdf");

        // Save the document as PDF.
        doc.Save(outputPath);

        // Verify that the file was created.
        if (File.Exists(outputPath))
        {
            Console.WriteLine("Document saved successfully to: " + outputPath);
        }
        else
        {
            Console.WriteLine("Failed to save the document.");
        }
    }
}
