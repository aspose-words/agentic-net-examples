using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

namespace AsposeWordsTableCloneExample
{
    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Add a paragraph before the table.
            builder.Writeln("Paragraph before the table.");

            // Build a simple 2x2 table.
            Table table = builder.StartTable();
            builder.InsertCell();
            builder.Write("Cell 1,1");
            builder.InsertCell();
            builder.Write("Cell 1,2");
            builder.EndRow();

            builder.InsertCell();
            builder.Write("Cell 2,1");
            builder.InsertCell();
            builder.Write("Cell 2,2");
            builder.EndRow();

            builder.EndTable();

            // Add a paragraph after the original table.
            builder.Writeln("Paragraph after the original table.");

            // Locate the original table node.
            Table originalTable = (Table)doc.GetChildNodes(NodeType.Table, true)[0];

            // Clone the table (deep clone).
            Table clonedTable = (Table)originalTable.Clone(true);

            // Locate the paragraph after which the cloned table will be inserted.
            // The body now contains: Paragraph (0), Table (1), Paragraph (2)
            Paragraph referenceParagraph = doc.FirstSection.Body.Paragraphs[2];

            // Insert the cloned table after the reference paragraph.
            referenceParagraph.ParentNode.InsertAfter(clonedTable, referenceParagraph);

            // Save the document.
            string outputPath = Path.Combine(Environment.CurrentDirectory, "TableCloneExample.docx");
            doc.Save(outputPath);

            // Simple validation to ensure the file was created.
            if (!File.Exists(outputPath))
                throw new InvalidOperationException("The output document was not saved correctly.");
        }
    }
}
