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

        // Build a sample 3x3 table using the DocumentBuilder workflow.
        Table table = builder.StartTable();

        for (int row = 0; row < 3; row++)
        {
            for (int col = 0; col < 3; col++)
            {
                builder.InsertCell();
                builder.Write($"R{row + 1}C{col + 1}");
            }
            builder.EndRow();
        }

        builder.EndTable();

        // Define the uniform width (in points) to apply to every column.
        double uniformWidth = 100.0;

        // Determine the number of columns from the first row.
        int columnCount = table.FirstRow.Cells.Count;

        // Iterate through each row and set the CellFormat.Width for each cell in the column.
        foreach (Row row in table.Rows)
        {
            for (int col = 0; col < columnCount; col++)
            {
                row.Cells[col].CellFormat.Width = uniformWidth;
            }
        }

        // Save the document to the current working directory.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "UniformColumnWidths.docx");
        doc.Save(outputPath);
    }
}
