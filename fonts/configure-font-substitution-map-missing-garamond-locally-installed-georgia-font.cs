using System;
using Aspose.Words;
using Aspose.Words.Fonts;

class FontSubstitutionExample
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Add some text that uses the missing font "Garamond".
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Font.Name = "Garamond";
        builder.Writeln("This text is formatted with Garamond, but will be rendered using Georgia.");

        // Configure font substitution: map "Garamond" to the locally installed "Georgia".
        FontSettings fontSettings = new FontSettings();
        fontSettings.SubstitutionSettings.TableSubstitution.SetSubstitutes("Garamond", "Georgia");

        // Apply the font settings to the document.
        doc.FontSettings = fontSettings;

        // Save the document (PDF format used as an example).
        doc.Save("FontSubstitutionOutput.pdf");
    }
}
