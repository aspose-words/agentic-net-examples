using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Drawing; // For Aspose.Drawing.Color

namespace FontFillExample
{
    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();

            // Use DocumentBuilder to add a paragraph with a single run of text.
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.Writeln("Sample text with red fill and 30% transparency.");

            // Retrieve the first run that was just added.
            Run run = doc.FirstSection.Body.FirstParagraph.Runs[0];

            // Access the Fill formatting of the run's font.
            Fill fill = run.Font.Fill;

            // Ensure the fill type is solid.
            fill.Solid();

            // Create a red color using Aspose.Drawing.Color.
            Aspose.Drawing.Color asposeDrawColor = Aspose.Drawing.Color.Red;

            // Convert Aspose.Drawing.Color to System.Drawing.Color and assign to the fill.
            // Fill.Color expects System.Drawing.Color, so we use System.Drawing.Color.FromArgb for conversion.
            fill.Color = System.Drawing.Color.FromArgb(asposeDrawColor.ToArgb());

            // Set the fill transparency to 30% (0.3 = 30% transparent, 0.7 opaque).
            fill.Transparency = 0.3;

            // Validation: output the assigned color ARGB and transparency to the console.
            Console.WriteLine($"Fill Color ARGB: {fill.Color.ToArgb()}");
            Console.WriteLine($"Fill Transparency: {fill.Transparency * 100}%");

            // Save the document to a file.
            string outputPath = "FontFillResult.docx";
            doc.Save(outputPath);

            // Verify that the file was created.
            Console.WriteLine($"Document saved: {File.Exists(outputPath)}");
        }
    }
}
