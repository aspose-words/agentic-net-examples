using System;
using System.IO;
using Aspose.Words;

namespace ParagraphSpacingExample
{
    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();

            // Initialize a DocumentBuilder for the document.
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Set spacing before and after for all paragraphs that will be created by this builder.
            // Values are in points (1 point = 1/72 inch). Here we use 12 points (1/6 inch).
            builder.ParagraphFormat.SpaceBefore = 12;
            builder.ParagraphFormat.SpaceAfter = 12;

            // Insert the first paragraph – it will inherit the spacing set above.
            builder.Writeln("First paragraph with 12 points spacing before and after.");

            // Insert a second paragraph with the same spacing.
            builder.Writeln("Second paragraph with the same spacing.");

            // Retrieve the second paragraph node to demonstrate modifying its spacing individually.
            Paragraph secondParagraph = doc.FirstSection.Body.Paragraphs[1];

            // Change spacing for this specific paragraph only.
            secondParagraph.ParagraphFormat.SpaceBefore = 24; // 24 points before
            secondParagraph.ParagraphFormat.SpaceAfter = 6;   // 6 points after

            // Save the document to the current directory.
            string outputPath = Path.Combine(Environment.CurrentDirectory, "ParagraphSpacing.docx");
            doc.Save(outputPath);
        }
    }
}
