using System;
using System.Globalization;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Fonts;
using Aspose.Words.Saving;

namespace HyperlinkLocalizationExample
{
    class Program
    {
        static void Main()
        {
            // Define the culture for which we want the localized hyperlink text.
            CultureInfo culture = new CultureInfo("fr-FR"); // Example: French (France)

            // Retrieve the localized display text without using external resource files.
            string displayText = GetLocalizedLinkText(culture);

            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Write some introductory text.
            builder.Writeln("For more information, please visit the following link:");

            // Apply hyperlink styling (blue color and single underline).
            builder.Font.Color = Color.Blue;
            builder.Font.Underline = Underline.Single;

            // Insert the hyperlink using the localized display text.
            builder.InsertHyperlink(displayText, "https://www.example.com", false);

            // Reset formatting to default for subsequent text.
            builder.Font.ClearFormatting();
            builder.Writeln();

            // Save the document to disk.
            doc.Save("LocalizedHyperlink.docx");
        }

        private static string GetLocalizedLinkText(CultureInfo culture)
        {
            // Simple in‑code localization fallback.
            return culture.Name switch
            {
                "fr-FR" => "Lien vers l'exemple",
                "es-ES" => "Enlace al ejemplo",
                "de-DE" => "Beispiel‑Link",
                _ => "Example Link"
            };
        }
    }
}
