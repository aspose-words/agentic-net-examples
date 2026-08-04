using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Lists;

namespace ListFormattingDemo
{
    public class Program
    {
        public static void Main()
        {
            // Create a new empty document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Build a numbered list with three items.
            builder.ListFormat.ApplyNumberDefault();
            builder.Writeln("First list item");
            builder.Writeln("Second list item");
            builder.Writeln("Third list item");

            // End the list so subsequent paragraphs are not automatically formatted.
            builder.ListFormat.RemoveNumbers();

            // Add a normal paragraph after the list.
            builder.Writeln("A normal paragraph without list formatting.");

            // Retrieve all paragraphs in the document.
            NodeCollection paragraphs = doc.GetChildNodes(NodeType.Paragraph, true);

            // Remove list formatting from the first two list items while keeping their indentation.
            var listParagraphs = paragraphs
                .OfType<Paragraph>()
                .Where(p => p.ListFormat.IsListItem)
                .Take(2);

            foreach (Paragraph para in listParagraphs)
            {
                // This call removes the bullet/number but leaves the paragraph's indentation intact.
                para.ListFormat.RemoveNumbers();
            }

            // Save the resulting document to the current directory.
            string outputPath = Path.Combine(Environment.CurrentDirectory, "ListFormattingResult.docx");
            doc.Save(outputPath);
        }
    }
}
