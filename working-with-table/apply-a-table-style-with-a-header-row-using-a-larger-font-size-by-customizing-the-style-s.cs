using System;
using System.Drawing;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

public class TableStyleExample
{
    public static void Main()
    {
        // Create a new document and a DocumentBuilder.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Create a custom table style.
        Style tableStyle = doc.Styles.Add(StyleType.Table, "MyCustomTableStyle");
        // Default font size for the whole table.
        tableStyle.Font.Size = 10;

        // Build the table.
        builder.StartTable();

        // Header row – use a larger font size.
        builder.Font.Size = 14; // Larger font for header cells.
        builder.InsertCell();
        builder.Writeln("Header 1");
        builder.InsertCell();
        builder.Writeln("Header 2");
        builder.EndRow();

        // Reset font size for data rows.
        builder.Font.Size = 10;

        // First data row.
        builder.InsertCell();
        builder.Writeln("Row 1, Cell 1");
        builder.InsertCell();
        builder.Writeln("Row 1, Cell 2");
        builder.EndRow();

        // Second data row.
        builder.InsertCell();
        builder.Writeln("Row 2, Cell 1");
        builder.InsertCell();
        builder.Writeln("Row 2, Cell 2");
        builder.EndRow();

        // End the table.
        builder.EndTable();

        // Retrieve the created table (the first table in the document).
        Table table = (Table)doc.GetChild(NodeType.Table, 0, true);

        // Apply the custom style to the table.
        table.StyleName = "MyCustomTableStyle";

        // Optional: give the header row a background color for visual distinction.
        Row headerRow = table.FirstRow;
        foreach (Cell cell in headerRow.Cells)
        {
            // Apply shading to each cell in the header row.
            cell.CellFormat.Shading.BackgroundPatternColor = Color.LightGray;
        }

        // Save the document.
        string outputPath = "TableStyleExample.docx";
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new Exception("The output document was not saved correctly.");
    }
}
