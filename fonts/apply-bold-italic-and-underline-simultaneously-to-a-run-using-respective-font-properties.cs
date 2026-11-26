using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fonts;
using Newtonsoft.Json; // Required package, kept for compliance

namespace FontFormattingExample
{
    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();

            // Create a paragraph to hold the run.
            Paragraph paragraph = new Paragraph(doc);

            // Create a run with sample text.
            Run run = new Run(doc, "Bold, Italic, and Underline");

            // Apply bold, italic, and underline formatting.
            run.Font.Bold = true;
            run.Font.Italic = true;
            run.Font.Underline = Aspose.Words.Underline.Single;

            // Validate that the properties were set correctly.
            bool isBold = run.Font.Bold;
            bool isItalic = run.Font.Italic;
            bool isUnderline = run.Font.Underline == Aspose.Words.Underline.Single;

            // Append the run to the paragraph and the paragraph to the document body.
            paragraph.AppendChild(run);
            doc.FirstSection.Body.AppendChild(paragraph);

            // Define output file path.
            string outputPath = "FormattedRun.docx";

            // Save the document.
            doc.Save(outputPath);

            // Verify that the file was created.
            if (File.Exists(outputPath))
            {
                Console.WriteLine("Document saved successfully.");
                Console.WriteLine($"Bold: {isBold}, Italic: {isItalic}, Underline: {isUnderline}");
            }
            else
            {
                Console.WriteLine("Failed to save the document.");
            }
        }
    }
}
