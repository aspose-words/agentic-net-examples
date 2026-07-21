using System;
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

        // Start the table.
        Table table = builder.StartTable();

        // ---------- Header row ----------
        // Apply a distinct background color for the header.
        builder.CellFormat.Shading.BackgroundPatternColor = Color.LightBlue;

        builder.InsertCell();
        builder.Writeln("Header 1");
        builder.InsertCell();
        builder.Writeln("Header 2");
        builder.EndRow();

        // ---------- Data rows ----------
        // Generate several data rows with alternating shading based on row index parity.
        int dataRowCount = 6; // Example number of data rows.
        for (int i = 0; i < dataRowCount; i++)
        {
            // Choose shading color: even rows white, odd rows light gray.
            Color rowColor = (i % 2 == 0) ? Color.White : Color.LightGray;
            builder.CellFormat.Shading.BackgroundPatternColor = rowColor;

            builder.InsertCell();
            builder.Writeln($"Row {i + 1} - Col 1");
            builder.InsertCell();
            builder.Writeln($"Row {i + 1} - Col 2");
            builder.EndRow();
        }

        // End the table.
        builder.EndTable();

        // Save the document to the local file system.
        string outputPath = "AlternatingRows.docx";
        doc.Save(outputPath);
    }
}
