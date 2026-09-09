using System;
using System.IO;
using Aspose.Words;

namespace ParagraphRangeExample
{
    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();

            // Use DocumentBuilder to add some paragraphs.
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.Writeln("This is the first paragraph.");
            builder.Writeln("This is the second paragraph.");

            // Retrieve the first paragraph from the document.
            Paragraph firstParagraph = doc.FirstSection.Body.Paragraphs[0];

            // Copy the paragraph's range content into a string variable.
            string paragraphContent = firstParagraph.Range.Text;

            // The range text includes the paragraph break character; trim if not needed.
            paragraphContent = paragraphContent.Trim();

            // Example usage of the extracted text (write to console).
            Console.WriteLine("Extracted paragraph text:");
            Console.WriteLine(paragraphContent);

            // Save the document to the local file system (optional verification).
            string outputPath = Path.Combine(Environment.CurrentDirectory, "SampleDocument.docx");
            doc.Save(outputPath);
        }
    }
}
