using System;
using Aspose.Words;

namespace RangeReplaceExample
{
    class Program
    {
        static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();

            // Add a paragraph with some initial text.
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.Writeln("This is the original paragraph text.");

            // Get the range of the first paragraph.
            Paragraph firstParagraph = doc.FirstSection.Body.FirstParagraph;
            Aspose.Words.Range paragraphRange = firstParagraph.Range;

            // Replace the paragraph's text using Range.Replace (Range.Text is read‑only).
            string originalText = paragraphRange.Text; // Includes the paragraph mark.
            paragraphRange.Replace(originalText, "This is the replaced paragraph text.");

            // Save the modified document.
            string outputPath = "Output.docx";
            doc.Save(outputPath);

            // Load the saved document to verify the replacement.
            Document loadedDoc = new Document(outputPath);
            string resultText = loadedDoc.FirstSection.Body.FirstParagraph.GetText();

            // Output the result to the console.
            Console.WriteLine("Resulting paragraph text:");
            Console.WriteLine(resultText);
        }
    }
}
