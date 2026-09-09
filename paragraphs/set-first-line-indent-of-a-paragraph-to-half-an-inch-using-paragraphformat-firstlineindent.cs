using System;
using Aspose.Words;

namespace ParagraphFirstLineIndentExample
{
    class Program
    {
        static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();

            // Initialize a DocumentBuilder for the document.
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Set the first line indent to half an inch (36 points).
            builder.ParagraphFormat.FirstLineIndent = 36.0;

            // Add a paragraph to demonstrate the indent.
            builder.Writeln("This paragraph has a first line indent of half an inch.");

            // Save the document.
            doc.Save("FirstLineIndent.docx");
        }
    }
}
