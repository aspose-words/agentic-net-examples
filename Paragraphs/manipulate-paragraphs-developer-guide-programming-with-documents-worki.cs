using System;
using Aspose.Words;

namespace AsposeWordsParagraphDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();

            // Initialize a DocumentBuilder which provides methods to insert content.
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert a paragraph with some text.
            // The Writeln method writes the text and then inserts a paragraph break.
            builder.Writeln("This is the inserted paragraph.");

            // Save the document to a DOCX file.
            doc.Save("InsertedParagraph.docx");
        }
    }
}
