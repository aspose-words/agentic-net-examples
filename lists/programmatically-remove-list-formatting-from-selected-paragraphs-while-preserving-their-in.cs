using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Lists;
using Aspose.Words.Tables;

namespace ListFormattingExample
{
    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Build a simple numbered list.
            builder.ListFormat.ApplyNumberDefault();
            builder.Writeln("First list item");
            builder.Writeln("Second list item");
            builder.Writeln("Third list item");

            // End the list.
            builder.ListFormat.RemoveNumbers();

            // Add a normal paragraph after the list.
            builder.Writeln("A regular paragraph.");

            // Retrieve all paragraphs in the document.
            NodeCollection paragraphs = doc.GetChildNodes(NodeType.Paragraph, true);

            // Remove list formatting from each list item while preserving its left indentation.
            foreach (Paragraph para in paragraphs.OfType<Paragraph>().Where(p => p.ListFormat.IsListItem))
            {
                // Store the current left indentation.
                double leftIndent = para.ParagraphFormat.LeftIndent;

                // Remove the list numbering/bullet.
                para.ListFormat.RemoveNumbers();

                // Reapply the stored indentation.
                para.ParagraphFormat.LeftIndent = leftIndent;
            }

            // Save the document to the current directory.
            string outputPath = Path.Combine(Environment.CurrentDirectory, "ListFormattingRemoved.docx");
            doc.Save(outputPath);
        }
    }
}
