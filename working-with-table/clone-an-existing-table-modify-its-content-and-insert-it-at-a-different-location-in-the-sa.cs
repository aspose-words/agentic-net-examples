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

            // Add some introductory text.
            builder.Writeln("Document with original table:");

            // Build the original table (2 rows x 2 columns).
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
            builder.EndRow();
            builder.EndTable();

            // Add a paragraph after the original table to serve as an insertion point.
            builder.Writeln("Paragraph after original table.");

            // Retrieve the first table from the document (the one we just created).
            Table firstTable = doc.FirstSection.Body.Tables[0];

            // Clone the table (deep clone, including its contents).
            Table clonedTable = (Table)firstTable.Clone(true);

            // Modify the cloned table's content.
            // Change text in the first cell.
            Cell firstCell = clonedTable.FirstRow.FirstCell;
            // Ensure the cell has at least one paragraph and one run.
            if (firstCell.FirstParagraph?.Runs.Count > 0)
            {
                firstCell.FirstParagraph.Runs[0].Text = "Cloned R1C1";
            }
            else
            {
                // If the cell is empty, add a new run with the desired text.
                firstCell.FirstParagraph?.AppendChild(new Run(doc, "Cloned R1C1"));
            }

            // Insert the cloned table after the paragraph that follows the original table.
            // Locate the paragraph after the original table.
            Node afterParagraph = firstTable.NextSibling;
            if (afterParagraph != null && afterParagraph.NodeType == NodeType.Paragraph)
            {
                // Insert the cloned table after this paragraph.
                afterParagraph.ParentNode.InsertAfter(clonedTable, afterParagraph);
            }
            else
            {
                // Fallback: append the cloned table at the end of the document body.
                doc.FirstSection.Body.AppendChild(clonedTable);
            }

            // Save the resulting document.
            string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "ClonedTable.docx");
            doc.Save(outputPath);

            // Simple validation to ensure the file was created.
            if (!File.Exists(outputPath))
                throw new InvalidOperationException("The output document was not saved correctly.");
        }
    }
}
