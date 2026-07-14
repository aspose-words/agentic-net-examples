using System;
using System.IO;
using System.Drawing;
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

        // ----- Header row -----
        // Apply a distinct shading color for the header.
        builder.CellFormat.Shading.BackgroundPatternColor = Color.LightGray;

        builder.InsertCell();
        builder.Write("Header 1");
        builder.InsertCell();
        builder.Write("Header 2");
        builder.InsertCell();
        builder.Write("Header 3");
        builder.EndRow();

        // ----- Data rows -----
        int dataRowCount = 6; // Number of data rows to create.

        for (int i = 0; i < dataRowCount; i++)
        {
            // Alternate shading based on row index parity.
            Color rowColor = (i % 2 == 0) ? Color.White : Color.LightBlue;
            builder.CellFormat.Shading.BackgroundPatternColor = rowColor;

            builder.InsertCell();
            builder.Write($"Row {i + 1} Col 1");
            builder.InsertCell();
            builder.Write($"Row {i + 1} Col 2");
            builder.InsertCell();
            builder.Write($"Row {i + 1} Col 3");
            builder.EndRow();
        }

        // End the table.
        builder.EndTable();

        // Save the document to the local file system.
        string outputPath = "AlternatingRowColors.docx";
        doc.Save(outputPath);

        // Simple validation to ensure the file was created.
        if (!File.Exists(outputPath))
            throw new Exception("The document was not saved correctly.");
    }
}
