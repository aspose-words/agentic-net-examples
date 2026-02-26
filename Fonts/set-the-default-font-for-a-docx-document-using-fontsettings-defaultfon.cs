using System;
using Aspose.Words;
using Aspose.Words.Fonts;

namespace SetDefaultFontExample
{
    class Program
    {
        static void Main(string[] args)
        {
            // Create a new blank document
            Document doc = new Document();

            // Create a FontSettings object and assign it to the document
            FontSettings fontSettings = new FontSettings();
            doc.FontSettings = fontSettings;

            // Set the default font substitution rule – this will be used as the document's default font
            fontSettings.SubstitutionSettings.DefaultFontSubstitution.DefaultFontName = "Courier New";

            // Save the document as DOCX
            doc.Save("DefaultFontDocument.docx");
        }
    }
}
