using System;
using Aspose.Words;

namespace ParagraphRangeExample
{
    public class Program
    {
        public static void Main()
        {
            // Create a new document and add paragraphs with sample text.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.Writeln("This is the first paragraph.");
            builder.Writeln("This is the second paragraph.");

            // Save the document locally.
            string filePath = "Sample.docx";
            doc.Save(filePath);

            // Load the document from the saved file.
            Document loadedDoc = new Document(filePath);

            // Retrieve the first paragraph in the document.
            Paragraph firstParagraph = loadedDoc.FirstSection.Body.Paragraphs[0];

            // Copy the paragraph's range content into a string variable.
            string paragraphText = firstParagraph.Range.Text;

            // Use the extracted text for further processing (here we simply output it).
            Console.WriteLine("Extracted paragraph text:");
            Console.WriteLine(paragraphText);
        }
    }
}
