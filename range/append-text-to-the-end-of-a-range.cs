using System;
using Aspose.Words;

namespace AppendTextRangeExample
{
    public class Program
    {
        public static void Main()
        {
            // Create a new document.
            Document doc = new Document();

            // Add a paragraph with initial text.
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.Writeln("This is the original content.");

            // Move the builder to the end of the document and append additional text.
            builder.MoveToDocumentEnd();
            builder.Write(" This is the appended text.");

            // Save the document.
            string outputPath = "AppendedDocument.docx";
            doc.Save(outputPath);

            // Load the saved document and output its text to verify the append operation.
            Document loadedDoc = new Document(outputPath);
            Console.WriteLine(loadedDoc.Range.Text);
        }
    }
}
