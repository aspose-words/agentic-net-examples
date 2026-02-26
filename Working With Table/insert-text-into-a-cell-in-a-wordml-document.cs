using System;
using Aspose.Words;
using Aspose.Words.Tables;

class Program
{
    static void Main()
    {
        // Load an existing WordML (or DOCX) document.
        Document doc = new Document("Input.docx");   // replace with the actual WORDML file path

        // Create a DocumentBuilder for editing the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Ensure the document contains at least one table.
        // If no table exists, create a simple 2x2 table for demonstration.
        if (doc.GetChildNodes(NodeType.Table, true).Count == 0)
        {
            builder.StartTable();
            builder.InsertCell();               // Row 0, Column 0
            builder.InsertCell();               // Row 0, Column 1
            builder.EndRow();
            builder.InsertCell();               // Row 1, Column 0
            builder.InsertCell();               // Row 1, Column 1
            builder.EndTable();
        }

        // Move the cursor to the target cell.
        // Parameters: tableIndex, rowIndex, columnIndex, characterIndex.
        // Here we target the cell at row 1, column 1 (second row, second column).
        builder.MoveToCell(tableIndex: 0, rowIndex: 1, columnIndex: 1, characterIndex: 0);

        // Insert the desired text into the current cell.
        builder.Write("Inserted text");

        // Save the modified document.
        doc.Save("Output.docx");   // change extension if you need WORDML output (e.g., .xml)
    }
}
