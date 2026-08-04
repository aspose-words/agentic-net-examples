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

        // Write some text using a font that may be missing (Garamond).
        builder.Font.Name = "Garamond";
        builder.Writeln("This paragraph is formatted with Garamond. If Garamond is not available, it should be rendered with Georgia.");

        // Configure font substitution: map missing Garamond to Georgia.
        FontSettings fontSettings = new FontSettings();
        fontSettings.SubstitutionSettings.TableSubstitution.SetSubstitutes("Garamond", new[] { "Georgia" });
        doc.FontSettings = fontSettings;

        // Save the document to PDF.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "FontSubstitutionExample.pdf");
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
