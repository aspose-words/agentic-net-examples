using System;
using Aspose.Words;
using Aspose.Words.Fonts;

namespace DocumentFormattingExample
{
    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();

            // Initialize DocumentBuilder for the document.
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Configure the font: bold and single underline.
            builder.Font.Bold = true;
            builder.Font.Underline = Underline.Single;

            // Insert the header text with the configured formatting.
            builder.Writeln("This is a Bold and Underlined Header");

            // Save the document to a file.
            string outputPath = "HeaderFormatted.docx";
            doc.Save(outputPath);
        }
    }
}
