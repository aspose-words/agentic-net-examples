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
        builder.Writeln("This text is formatted with Garamond, which should be substituted with Georgia.");

        // Set up font substitution: map missing Garamond to the locally installed Georgia font.
        FontSettings fontSettings = new FontSettings();
        fontSettings.SubstitutionSettings.TableSubstitution.SetSubstitutes("Garamond", new[] { "Georgia" });

        // Apply the font settings to the document.
        doc.FontSettings = fontSettings;

        // Save the document to a PDF file in the current directory.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "FontSubstitutionExample.pdf");
        doc.Save(outputPath);
    }
}
