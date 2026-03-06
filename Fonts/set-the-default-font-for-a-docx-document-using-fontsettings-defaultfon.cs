using System;
using Aspose.Words;
using Aspose.Words.Fonts;

class SetDefaultFontExample
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Set the default font that will be used when a requested font is missing.
        // This modifies the static default FontSettings instance, which is applied to all documents
        // that do not have their own FontSettings assigned.
        FontSettings.DefaultInstance.SubstitutionSettings
            .DefaultFontSubstitution.DefaultFontName = "Courier New";

        // Optionally, also set the document's style collection default font
        // so that newly created styles inherit this font.
        doc.Styles.DefaultFont.Name = "Courier New";

        // Save the document to a DOCX file.
        doc.Save("DefaultFont.docx");
    }
}
