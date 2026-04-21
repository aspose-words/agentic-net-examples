using System;
using System.Drawing;
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

        // ----- Header row -----
        // Insert first header cell.
        builder.InsertCell();
        // Apply light gray shading to this cell.
        builder.CellFormat.Shading.BackgroundPatternColor = Color.LightGray;
        builder.Write("Header 1");

        // Insert second header cell.
        builder.InsertCell();
        // Apply the same shading to the second cell.
        builder.CellFormat.Shading.BackgroundPatternColor = Color.LightGray;
        builder.Write("Header 2");

        // End the header row.
        builder.EndRow();

        // ----- Data rows (example) -----
        // First data row.
        builder.InsertCell();
        builder.CellFormat.Shading.ClearFormatting(); // Remove shading for data cells.
        builder.Write("Row 1, Col 1");
        builder.InsertCell();
        builder.Write("Row 1, Col 2");
        builder.EndRow();

        // Second data row.
        builder.InsertCell();
        builder.Write("Row 2, Col 1");
        builder.InsertCell();
        builder.Write("Row 2, Col 2");
        builder.EndRow();

        // End the table.
        builder.EndTable();

        // Save the document to a file.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "HeaderRowShading.docx");
        doc.Save(outputPath);

        // Simple validation to ensure the file was created.
        if (!File.Exists(outputPath))
            throw new Exception("The output document was not created.");
    }
}
