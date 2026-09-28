using System;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a new document and a DocumentBuilder.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Define a custom table style programmatically.
        Style customTableStyle = doc.Styles.Add(StyleType.Table, "MyCustomTableStyle");
        customTableStyle.Font.Name = "Arial";
        customTableStyle.Font.Size = 10;

        // Start building the table.
        builder.StartTable();

        // Header row.
        builder.InsertCell();
        builder.Write("Header 1");
        builder.InsertCell();
        builder.Write("Header 2");
        builder.EndRow();

        // Add several data rows.
        for (int i = 1; i <= 5; i++)
        {
            builder.InsertCell();
            builder.Write($"Row {i} Col 1");
            builder.InsertCell();
            builder.Write($"Row {i} Col 2");
            builder.EndRow();
        }

        // End the table.
        builder.EndTable();

        // Retrieve the created table.
        Table table = (Table)doc.GetChild(NodeType.Table, 0, true);

        // Apply the custom style to the table.
        table.Style = customTableStyle;

        // Apply a simple border to the whole table.
        table.SetBorders(LineStyle.Single, 0.5, Color.Black);

        // Apply alternating row shading (skip header row).
        for (int rowIndex = 1; rowIndex < table.Rows.Count; rowIndex++)
        {
            Row row = table.Rows[rowIndex];
            Color shadeColor = (rowIndex % 2 == 0) ? Color.LightGray : Color.White;

            foreach (Cell cell in row.Cells)
            {
                cell.CellFormat.Shading.BackgroundPatternColor = shadeColor;
            }
        }

        // Save the document.
        string outputPath = "AlternatingRows.docx";
        doc.Save(outputPath);

        // Verify that the file was created (non‑interactive).
        if (System.IO.File.Exists(outputPath))
        {
            Console.WriteLine($"Document saved successfully to '{outputPath}'.");
        }
    }
}
