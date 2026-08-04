using System;
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

            // Add a paragraph that will serve as the insertion point.
            builder.Writeln("Paragraph before the original table.");

            // Build the original table.
            Table originalTable = builder.StartTable();
            builder.InsertCell();
            builder.Write("Cell 1");
            builder.InsertCell();
            builder.Write("Cell 2");
            builder.EndRow();
            builder.EndTable();

            // Add another paragraph after the original table (optional, just to have more content).
            builder.Writeln("Paragraph after the original table.");

            // Retrieve the paragraph after which we want to insert the cloned table.
            // In this example we use the first paragraph ("Paragraph before the original table.").
            Paragraph referenceParagraph = (Paragraph)doc.GetChild(NodeType.Paragraph, 0, true);

            // Clone the original table (deep clone).
            Table clonedTable = (Table)originalTable.Clone(true);

            // Insert the cloned table immediately after the reference paragraph.
            referenceParagraph.ParentNode.InsertAfter(clonedTable, referenceParagraph);

            // Simple validation: the document should now contain two tables.
            int tableCount = doc.GetChildNodes(NodeType.Table, true).Count;
            if (tableCount != 2)
                throw new InvalidOperationException($"Expected 2 tables, but found {tableCount}.");

            // Save the resulting document.
            doc.Save("ClonedTable.docx");
        }
    }
}
