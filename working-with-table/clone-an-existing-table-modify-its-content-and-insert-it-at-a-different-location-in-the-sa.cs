using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

namespace TableCloneExample
{
    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Build an original table with 2 rows and 2 columns.
            builder.StartTable();
            // First row
            builder.InsertCell();
            builder.Write("R1C1");
            builder.InsertCell();
            builder.Write("R1C2");
            builder.EndRow();
            // Second row
            builder.InsertCell();
            builder.Write("R2C1");
            builder.InsertCell();
            builder.Write("R2C2");
            builder.EndRow();
            builder.EndTable();

            // Retrieve the first (and only) table in the document.
            Table originalTable = (Table)doc.GetChildNodes(NodeType.Table, true)[0];

            // Clone the table (deep clone).
            Table clonedTable = (Table)originalTable.Clone(true);

            // Modify content of the cloned table: change text in the first cell.
            Cell firstCell = clonedTable.Rows[0].Cells[0];
            // Clear existing runs and add new text.
            firstCell.FirstParagraph.Runs.Clear();
            firstCell.FirstParagraph.AppendChild(new Run(doc, "Modified Cell"));

            // Insert the cloned table after the original table.
            // The parent of a Table is a CompositeNode (e.g., Body), which provides InsertAfter.
            CompositeNode parent = originalTable.ParentNode as CompositeNode;
            if (parent == null)
                throw new InvalidOperationException("Unable to locate a valid parent node for insertion.");

            parent.InsertAfter(clonedTable, originalTable);

            // Save the resulting document.
            string outputPath = "ClonedTable.docx";
            doc.Save(outputPath);

            // Simple validation that the file was created.
            if (!File.Exists(outputPath))
                throw new Exception("The output document was not created.");
        }
    }
}
