using System;
using Aspose.Words;
using Aspose.Words.Lists;

namespace ListFormattingExample
{
    class Program
    {
        static void Main()
        {
            // Create a new document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Start a numbered list.
            builder.ListFormat.ApplyNumberDefault();

            // Add list items.
            builder.Writeln("First item");
            builder.Writeln("Second item");
            builder.Writeln("Third item");

            // End list formatting for any following paragraphs.
            builder.ListFormat.RemoveNumbers();

            // Remove list formatting from the first two paragraphs while preserving indentation.
            Paragraph para1 = doc.FirstSection.Body.Paragraphs[0];
            Paragraph para2 = doc.FirstSection.Body.Paragraphs[1];

            para1.ListFormat.RemoveNumbers();
            para2.ListFormat.RemoveNumbers();

            // Save the document.
            string outputPath = "Result.docx";
            doc.Save(outputPath);
        }
    }
}
