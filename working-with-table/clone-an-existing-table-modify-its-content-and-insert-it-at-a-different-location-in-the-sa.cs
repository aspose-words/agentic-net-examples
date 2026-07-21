using System;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add a paragraph before the first table.
        builder.Writeln("Document with original table:");

        // Build the original table (2 rows x 2 columns).
        Table originalTable = builder.StartTable();
        builder.InsertCell();
        builder.Write("A1");
        builder.InsertCell();
        builder.Write("B1");
        builder.EndRow();
        builder.InsertCell();
        builder.Write("A2");
        builder.InsertCell();
        builder.Write("B2");
        builder.EndTable();

        // Add a paragraph after the original table to separate content.
        builder.Writeln("\nDocument after original table:");

        // Clone the original table (deep clone).
        Table clonedTable = (Table)originalTable.Clone(true);

        // Modify the cloned table's cell contents.
        int rowIndex = 0;
        foreach (Row row in clonedTable.Rows)
        {
            int cellIndex = 0;
            foreach (Cell cell in row.Cells)
            {
                // Remove existing content.
                cell.RemoveAllChildren();

                // Add new paragraph with modified text.
                Paragraph para = new Paragraph(doc);
                Run run = new Run(doc, $"Cloned {rowIndex + 1},{cellIndex + 1}");
                para.AppendChild(run);
                cell.AppendChild(para);

                cellIndex++;
            }
            rowIndex++;
        }

        // Insert the cloned table after the original table.
        // The parent of a table is a CompositeNode (e.g., Body), so cast accordingly.
        CompositeNode parent = originalTable.ParentNode as CompositeNode;
        if (parent == null)
            throw new InvalidOperationException("Unable to locate a valid parent node for insertion.");

        parent.InsertAfter(clonedTable, originalTable);

        // Save the resulting document.
        const string outputPath = "ClonedTable.docx";
        doc.Save(outputPath);
    }
}
