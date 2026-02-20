using System;
using Aspose.Words;
using Aspose.Words.Fonts;

class SetDefaultFontExample
{
    static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();

        // Initialize FontSettings and assign it to the document.
        FontSettings fontSettings = new FontSettings();
        doc.FontSettings = fontSettings;

        // Access the default font substitution rule.
        DefaultFontSubstitutionRule defaultFontRule = 
            fontSettings.SubstitutionSettings.DefaultFontSubstitution;

        // Set the default font name that will be used when a requested font is missing.
        defaultFontRule.DefaultFontName = "Arial";

        // Optional: demonstrate substitution by using a font that does not exist.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Font.Name = "Missing Font";
        builder.Writeln("This line uses a missing font and will be rendered with Arial.");

        // Save the document as DOCX.
        doc.Save("DefaultFontExample.docx");
    }
}
