using System;
using System.Globalization;
using System.Resources;
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
            // This could be changed at runtime based on user preferences.
            CultureInfo culture = new CultureInfo("fr-FR"); // Example: French (France)

            // Load the resource manager for the .resx file named "Resources.resx"
            // The base name is the fully qualified name of the resource file without the .resx extension.
            ResourceManager rm = new ResourceManager("HyperlinkLocalizationExample.Resources", typeof(Program).Assembly);

            // Retrieve the localized display text for the hyperlink.
            // The resource file must contain a string entry with the key "LinkText".
            string displayText = rm.GetString("LinkText", culture) ?? "Default Link";

            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Write some introductory text.
            builder.Writeln("For more information, please visit the following link:");

            // Apply hyperlink styling (blue color and single underline).
            builder.Font.Color = Color.Blue;
            builder.Font.Underline = Underline.Single;

            // Insert the hyperlink using the localized display text.
            // The third parameter 'false' indicates that the second argument is a URL, not a bookmark.
            builder.InsertHyperlink(displayText, "https://www.example.com", false);

            // Reset formatting to default for subsequent text.
            builder.Font.ClearFormatting();
            builder.Writeln();

            // Save the document to disk.
            doc.Save("LocalizedHyperlink.docx");
        }
    }
}
