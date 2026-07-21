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
            // Define a folder to store the output document.
            string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
            Directory.CreateDirectory(artifactsDir);

            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Start a default numbered list and add a few items.
            builder.ListFormat.ApplyNumberDefault();
            builder.Writeln("Numbered list item 1");
            builder.Writeln("Numbered list item 2");
            builder.Writeln("Numbered list item 3");

            // Retrieve all paragraphs in the document.
            NodeCollection paragraphs = doc.GetChildNodes(NodeType.Paragraph, true);

            // For each paragraph that is part of a list, remove its list formatting.
            foreach (Paragraph para in paragraphs)
            {
                if (para.ListFormat.IsListItem)
                {
                    para.ListFormat.RemoveNumbers();
                }
            }

            // Save the resulting document. The list items are now plain paragraphs.
            string outputPath = Path.Combine(artifactsDir, "ListWithoutNumbers.docx");
            doc.Save(outputPath);
        }
    }
}
