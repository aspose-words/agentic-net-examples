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

        // Start a table.
        Table table = builder.StartTable();

        // Ensure the table has at least one row/cell so that padding properties can be set.
        table.EnsureMinimum();

        // Set the default cell margins (padding) for the entire table (2 points on each side).
        table.TopPadding = 2;
        table.BottomPadding = 2;
        table.LeftPadding = 2;
        table.RightPadding = 2;

        // Add a simple row with two cells.
        builder.InsertCell();
        builder.Write("Cell 1");
        builder.InsertCell();
        builder.Write("Cell 2");
        builder.EndRow();

        // Finish the table.
        builder.EndTable();

        // Save the document.
        string outputPath = "Table.DefaultCellMargin.docx";
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new Exception("The output file was not created.");
    }
}
