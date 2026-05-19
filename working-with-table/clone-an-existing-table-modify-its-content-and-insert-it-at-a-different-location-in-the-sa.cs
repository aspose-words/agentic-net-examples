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

            // Build the original table with two rows and two columns.
            Table originalTable = builder.StartTable();
            builder.InsertCell();
            builder.Write("Original R1C1");
            builder.InsertCell();
            builder.Write("Original R1C2");
            builder.EndRow();

            builder.InsertCell();
            builder.Write("Original R2C1");
            builder.InsertCell();
            builder.Write("Original R2C2");
            builder.EndTable();

            // Append the table to the document body.
            doc.FirstSection.Body.AppendChild(originalTable);

            // Deep clone the original table.
            Table clonedTable = (Table)originalTable.Clone(true);

            // Modify the cloned table's first cell.
            Cell firstCell = clonedTable.FirstRow.FirstCell;
            firstCell.FirstParagraph.Runs.Clear(); // Remove existing runs.
            firstCell.FirstParagraph.AppendChild(new Run(doc, "Cloned R1C1"));

            // Insert the cloned table after the original table.
            // The parent of a table is a Body node, which derives from CompositeNode.
            ((CompositeNode)originalTable.ParentNode).InsertAfter(clonedTable, originalTable);

            // Save the document.
            string outputPath = "ClonedTable.docx";
            doc.Save(outputPath);

            // Verify that the file was created.
            if (!File.Exists(outputPath))
                throw new InvalidOperationException("The output document was not saved correctly.");
        }
    }
}
