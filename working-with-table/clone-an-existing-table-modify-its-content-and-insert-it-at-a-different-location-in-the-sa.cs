using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add an introductory paragraph.
        builder.Writeln("Document start.");

        // Build a sample table with two columns and three rows.
        builder.StartTable();

        // Header row.
        builder.InsertCell();
        builder.Write("Header 1");
        builder.InsertCell();
        builder.Write("Header 2");
        builder.EndRow();

        // First data row.
        builder.InsertCell();
        builder.Write("Row 1, Cell 1");
        builder.InsertCell();
        builder.Write("Row 1, Cell 2");
        builder.EndRow();

        // Second data row.
        builder.InsertCell();
        builder.Write("Row 2, Cell 1");
        builder.InsertCell();
        builder.Write("Row 2, Cell 2");
        builder.EndRow();

        // End the table.
        builder.EndTable();

        // Add a paragraph after the original table.
        builder.Writeln("Paragraph after original table.");

        // Locate the original table (the first table in the document).
        Table originalTable = (Table)doc.GetChild(NodeType.Table, 0, true);
        if (originalTable == null)
            throw new Exception("Original table not found.");

        // Clone the table deeply.
        Table clonedTable = (Table)originalTable.Clone(true);

        // Modify the content of the cloned table: change text in the first cell.
        Cell firstCell = clonedTable.FirstRow.FirstCell;
        firstCell.RemoveAllChildren(); // Remove existing paragraphs.
        Paragraph newParagraph = new Paragraph(doc);
        Run newRun = new Run(doc, "Modified Header 1");
        newParagraph.AppendChild(newRun);
        firstCell.AppendChild(newParagraph);

        // Insert the cloned table after the original table.
        // The parent of a Table is a CompositeNode (e.g., Body). Cast to CompositeNode to use InsertAfter.
        CompositeNode parent = (CompositeNode)originalTable.ParentNode;
        parent.InsertAfter(clonedTable, originalTable);

        // Save the resulting document.
        string outputPath = "ClonedTable.docx";
        doc.Save(outputPath);

        // Verify that the file was saved.
        if (!File.Exists(outputPath))
            throw new Exception("Failed to save the document.");

        // Indicate successful completion.
        Console.WriteLine("Document saved to " + Path.GetFullPath(outputPath));
    }
}
