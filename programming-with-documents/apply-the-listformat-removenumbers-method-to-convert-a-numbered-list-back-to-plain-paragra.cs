using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Lists;

namespace ListRemoveNumbersExample
{
    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();

            // Use DocumentBuilder to add a numbered list.
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.ListFormat.ApplyNumberDefault(); // Start a default numbered list.
            builder.Writeln("Numbered item 1");
            builder.Writeln("Numbered item 2");
            builder.Writeln("Numbered item 3");

            // Convert each list item back to a plain paragraph by removing list formatting.
            NodeCollection paragraphs = doc.GetChildNodes(NodeType.Paragraph, true);
            foreach (Paragraph paragraph in paragraphs)
            {
                // Remove numbers or bullets from the paragraph.
                paragraph.ListFormat.RemoveNumbers();
            }

            // Save the resulting document.
            string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "ListRemoved.docx");
            doc.Save(outputPath);
        }
    }
}
