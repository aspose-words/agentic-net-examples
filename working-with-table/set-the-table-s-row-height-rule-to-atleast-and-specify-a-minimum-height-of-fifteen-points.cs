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

        // Start a new table.
        Table table = builder.StartTable();

        // Configure the current row to have a minimum height of 15 points.
        // HeightRule.AtLeast ensures the row will be at least this height.
        builder.RowFormat.Height = 15;
        builder.RowFormat.HeightRule = HeightRule.AtLeast;

        // Add a cell to the first row and write some text.
        builder.InsertCell();
        builder.Write("Row with AtLeast height.");

        // Finish the first row.
        builder.EndRow();

        // Add a second row with default formatting.
        builder.InsertCell();
        builder.Write("Second row.");
        builder.EndRow();

        // End the table.
        builder.EndTable();

        // Save the document to a file.
        doc.Save("TableRowHeight.docx");
    }
}
