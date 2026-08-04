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

            // Build the original table (2 rows x 2 columns) using the builder.
            Table originalTable = builder.StartTable();
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
            builder.EndTable(); // The table is now part of the document.

            // Clone the existing table (deep clone).
            Table clonedTable = (Table)originalTable.Clone(true);

            // Modify the cloned table's content.
            foreach (Row row in clonedTable.Rows)
            {
                foreach (Cell cell in row.Cells)
                {
                    // Each cell already contains a paragraph with a run.
                    // Replace the text of the first run with new content.
                    if (cell.FirstParagraph != null && cell.FirstParagraph.Runs.Count > 0)
                    {
                        cell.FirstParagraph.Runs[0].Text = "Cloned";
                    }
                    else
                    {
                        // Ensure the cell has a paragraph and add a run if needed.
                        cell.EnsureMinimum();
                        cell.FirstParagraph.AppendChild(new Run(doc, "Cloned"));
                    }
                }
            }

            // Insert a marker paragraph where the cloned table will be placed.
            builder.MoveToDocumentEnd();
            builder.Writeln("=== Cloned Table Inserted Below ===");
            // Capture the marker paragraph node.
            Paragraph markerParagraph = (Paragraph)doc.LastSection.Body.LastParagraph;

            // Insert the cloned table after the marker paragraph.
            // Use InsertAfter on the parent node (the body) to place the table correctly.
            markerParagraph.ParentNode.InsertAfter(clonedTable, markerParagraph);

            // Save the resulting document.
            string outputPath = "ClonedTable.docx";
            doc.Save(outputPath);

            // Simple validation to ensure the file was created.
            if (!File.Exists(outputPath))
                throw new Exception("The output document was not created.");
        }
    }
}
