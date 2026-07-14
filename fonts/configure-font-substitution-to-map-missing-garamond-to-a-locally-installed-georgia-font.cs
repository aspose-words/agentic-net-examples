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

        // Write some text using a font that is likely missing on the system (Garamond).
        builder.Font.Name = "Garamond";
        builder.Writeln("This text is formatted with Garamond, which will be substituted.");

        // Configure font substitution: map missing "Garamond" to the locally installed "Georgia".
        FontSettings fontSettings = new FontSettings();
        // Use the table substitution rule to specify the substitute font.
        fontSettings.SubstitutionSettings.TableSubstitution.SetSubstitutes(
            "Garamond", new[] { "Georgia" });

        // Assign the configured FontSettings to the document.
        doc.FontSettings = fontSettings;

        // Save the document to PDF. The missing Garamond font will be rendered using Georgia.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "FontSubstitutionExample.pdf");
        doc.Save(outputPath);
    }
}
