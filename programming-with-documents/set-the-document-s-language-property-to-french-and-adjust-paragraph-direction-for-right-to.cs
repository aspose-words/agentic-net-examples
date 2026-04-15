using System;
using System.Globalization;
using System.IO;
using Aspose.Words;

namespace AsposeWordsDemo
{
    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();

            // Set the document's default language to French (fr-FR, LCID = 1036).
            // This will affect the language of newly added text unless overridden.
            doc.Styles.DefaultFont.LocaleId = new CultureInfo("fr-FR").LCID;

            // Use DocumentBuilder to add content.
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Write a French sentence.
            builder.Writeln("Bonjour le monde! Ceci est un texte en français.");

            // Adjust paragraph direction for right‑to‑left scripts.
            // Mark the current paragraph as RTL.
            builder.CurrentParagraph.ParagraphFormat.Bidi = true;

            // Add a sample RTL paragraph (e.g., Hebrew) to demonstrate the direction.
            builder.Writeln("שלום עולם! זה טקסט מימין לשמאל.");

            // Define the output path.
            string outputPath = Path.Combine(Environment.CurrentDirectory, "Result.docx");

            // Save the document.
            doc.Save(outputPath);
        }
    }
}
