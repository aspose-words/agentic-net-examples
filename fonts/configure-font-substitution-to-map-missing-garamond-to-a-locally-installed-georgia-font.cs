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

        // Write a paragraph using a font that might be missing on the system.
        builder.Font.Name = "Garamond";
        builder.Writeln("This text is formatted with Garamond, which may be unavailable.");

        // Configure font substitution: map the missing Garamond font to the locally installed Georgia font.
        FontSettings fontSettings = new FontSettings();
        fontSettings.SubstitutionSettings.TableSubstitution.SetSubstitutes(
            "Garamond", new[] { "Georgia" });

        // Apply the font settings to the document.
        doc.FontSettings = fontSettings;

        // Save the document to PDF. The text will be rendered using Georgia if Garamond is not found.
        string outputPath = "FontSubstitutionExample.pdf";
        doc.Save(outputPath);

        // Ensure the file was created (no console output as per requirements).
        if (File.Exists(outputPath))
        {
            // File exists – nothing else to do.
        }
    }
}
