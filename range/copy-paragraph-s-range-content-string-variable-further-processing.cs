using System;
using Aspose.Words;

namespace ParagraphRangeExample
{
    class Program
    {
        static void Main()
        {
            // Load an existing Word document.
            Document doc = new Document("Input.docx");

            // Get the first paragraph in the document's main body.
            Paragraph paragraph = doc.FirstSection.Body.Paragraphs[0];

            // Copy the paragraph's range content into a string variable.
            // The Range.Text property returns the text covered by the paragraph,
            // including any control characters (e.g., paragraph break).
            string paragraphContent = paragraph.Range.Text;

            // Example of further processing: trim trailing control characters.
            string processedContent = paragraphContent.Trim();

            // Output the extracted text to the console.
            Console.WriteLine("Original paragraph text:");
            Console.WriteLine(paragraphContent);
            Console.WriteLine("\nProcessed paragraph text:");
            Console.WriteLine(processedContent);
        }
    }
}
