using System;
using System.IO;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a new document and a builder for it.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Start a table.
        Table table = builder.StartTable();

        // ----- Header row -----
        builder.InsertCell();
        builder.Write("Header 1");
        builder.InsertCell();
        builder.Write("Header 2");
        builder.InsertCell();
        builder.Write("Header 3");
        builder.EndRow();

        // ----- Data rows -----
        for (int i = 1; i <= 3; i++)
        {
            builder.InsertCell();
            builder.Write($"Row {i} Col 1");
            builder.InsertCell();
            builder.Write($"Row {i} Col 2");
            builder.InsertCell();
            builder.Write($"Row {i} Col 3");
            builder.EndRow();
        }

        // Finish the table.
        builder.EndTable();

        // Apply light gray shading to each cell in the header row.
        foreach (Cell cell in table.FirstRow.Cells)
        {
            cell.CellFormat.Shading.BackgroundPatternColor = Color.LightGray;
        }

        // Save the document.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "HeaderRowShading.docx");
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
        {
            throw new Exception("The document was not saved correctly.");
        }
    }
}
