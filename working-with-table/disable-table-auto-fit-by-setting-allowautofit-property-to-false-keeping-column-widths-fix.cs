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

        // Start a new table.
        builder.StartTable();

        // First column – set a fixed width of 100 points.
        builder.CellFormat.Width = 100;
        builder.InsertCell();
        builder.Writeln("First column");

        // Second column – set a fixed width of 200 points.
        builder.CellFormat.Width = 200;
        builder.InsertCell();
        builder.Writeln("Second column");

        // End the row and the table.
        builder.EndRow();
        builder.EndTable();

        // Retrieve the created table and disable auto‑fit.
        Table table = (Table)doc.GetChild(NodeType.Table, 0, true);
        table.AllowAutoFit = false;

        // Save the document.
        string outputPath = "TableAutoFit.docx";
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new Exception("The output file was not created.");

        // Optionally, you could open the document to confirm settings,
        // but the task requires only creation and saving.
    }
}
