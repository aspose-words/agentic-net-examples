using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Tables;

namespace AsposeWordsTableInsertExample
{
    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Add a heading paragraph that will be the anchor for the table.
            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
            builder.Writeln("Sample Heading");

            // Add another paragraph just to have more content after the heading.
            builder.ParagraphFormat.ClearFormatting();
            builder.Writeln("Paragraph after the heading.");

            // Build a simple 2‑column table at the end of the document.
            builder.MoveToDocumentEnd();
            Table table = builder.StartTable();
            builder.InsertCell();
            builder.Write("Cell 1");
            builder.InsertCell();
            builder.Write("Cell 2");
            builder.EndRow();
            builder.EndTable(); // The table is now part of the document at the end.

            // Locate the heading paragraph we created earlier.
            Paragraph heading = doc.GetChildNodes(NodeType.Paragraph, true)
                                   .Cast<Paragraph>()
                                   .FirstOrDefault(p => p.ParagraphFormat.StyleIdentifier == StyleIdentifier.Heading1);

            if (heading == null)
                throw new InvalidOperationException("Heading paragraph not found.");

            // Insert the table after the heading paragraph.
            // InsertAfter must be called on the parent node of the reference node.
            heading.ParentNode.InsertAfter(table, heading);

            // Save the document.
            string outputPath = Path.Combine(Environment.CurrentDirectory, "OutputTableAfterHeading.docx");
            doc.Save(outputPath);

            // Verify that the file was created.
            if (!File.Exists(outputPath))
                throw new FileNotFoundException("The output document was not saved.", outputPath);
        }
    }
}
