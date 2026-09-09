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
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Write a line using a font that may not be installed (Garamond).
        builder.Font.Name = "Garamond";
        builder.Writeln("This text is formatted with Garamond, which may be missing on the system.");

        // Set up font substitution: replace missing Garamond with Georgia.
        FontSettings fontSettings = new FontSettings();
        fontSettings.SubstitutionSettings.TableSubstitution.SetSubstitutes(
            "Garamond", new[] { "Georgia" });

        // Apply the font settings to the document.
        doc.FontSettings = fontSettings;

        // Save the document to PDF.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "FontSubstitutionExample.pdf");
        doc.Save(outputPath);

        // Verify that the file was created.
        if (File.Exists(outputPath))
        {
            Console.WriteLine("Document saved successfully: " + outputPath);
        }
        else
        {
            Console.WriteLine("Failed to save the document.");
        }
    }
}
