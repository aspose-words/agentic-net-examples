using System;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // ---------- Build the outer table ----------
        builder.StartTable();

        // First row
        builder.InsertCell();
        builder.Write("Outer 1,1");
        builder.InsertCell();
        builder.Write("Outer 1,2");
        builder.EndRow();

        // Second row
        builder.InsertCell();
        builder.Write("Outer 2,1");
        builder.InsertCell();
        builder.Write("Outer 2,2");
        Table outerTable = builder.EndTable();

        // Get a reference to the first cell of the outer table.
        Cell targetCell = outerTable.FirstRow.FirstCell;

        // ---------- Build the inner (nested) table ----------
        // Move the builder's cursor into the target cell.
        builder.MoveTo(targetCell.FirstParagraph);

        builder.StartTable();

        // First row of inner table
        builder.InsertCell();
        builder.Write("Inner 1,1");
        builder.InsertCell();
        builder.Write("Inner 1,2");
        builder.EndRow();

        // Second row of inner table
        builder.InsertCell();
        builder.Write("Inner 2,1");
        builder.InsertCell();
        builder.Write("Inner 2,2");
        Table innerTable = builder.EndTable();

        // Append the inner table to the target cell.
        targetCell.AppendChild(innerTable);

        // Save the document.
        const string outputPath = "NestedTable.docx";
        doc.Save(outputPath);
    }
}
