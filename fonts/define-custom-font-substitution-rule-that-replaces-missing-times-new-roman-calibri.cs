// Program.cs
using System;
using Aspose.Words;
using Aspose.Words.Fonts;

class Program
{
    static void Main()
    {
        // Create a new blank document
        Document doc = new Document();

        // Initialize font settings for the document
        FontSettings fontSettings = new FontSettings();
        doc.FontSettings = fontSettings;

        // Configure a custom substitution: if "Times New Roman" is missing, use "Calibri"
        // The second argument can be a params array of substitute font names.
        fontSettings.SubstitutionSettings.TableSubstitution.AddSubstitutes(
            "Times New Roman", "Calibri");

        // Add some text using the original font to demonstrate the substitution
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Font.Name = "Times New Roman";
        builder.Writeln("This line is formatted with Times New Roman, but will be rendered with Calibri if the original font is unavailable.");

        // Save the document (PDF format used as an example)
        doc.Save("CustomFontSubstitution.pdf");
    }
}
