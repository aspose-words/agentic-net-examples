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

        // Start building a table.
        builder.StartTable();

        // Insert 3 rows with 3 cells each.
        for (int row = 1; row <= 3; row++)
        {
            for (int col = 1; col <= 3; col++)
            {
                builder.InsertCell();
                builder.Writeln($"Row {row}, Cell {col}");
            }
            builder.EndRow();
        }

        // End the table.
        builder.EndTable();

        // Retrieve the created table.
        Table table = (Table)doc.GetChild(NodeType.Table, 0, true);
        if (table == null)
            throw new InvalidOperationException("Table was not created.");

        // Apply the built‑in style "Grid Table 5 Dark".
        // The style can be obtained from the document's style collection by name.
        Style gridStyle = doc.Styles["Grid Table 5 Dark"];
        if (gridStyle == null)
            throw new InvalidOperationException("The style 'Grid Table 5 Dark' was not found.");
        table.Style = gridStyle;

        // Save the document.
        string outputPath = "StyledTable.docx";
        doc.Save(outputPath);

        // Verify that the file was saved.
        if (!File.Exists(outputPath))
            throw new FileNotFoundException("The output document was not saved.", outputPath);
    }
}
