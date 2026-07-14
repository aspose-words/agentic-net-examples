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

        // Build a sample table (3 rows x 4 columns).
        Table table = builder.StartTable();

        int rows = 3;
        int cols = 4;
        for (int r = 0; r < rows; r++)
        {
            for (int c = 0; c < cols; c++)
            {
                builder.InsertCell();
                builder.Write($"R{r + 1}C{c + 1}");
            }
            builder.EndRow();
        }

        builder.EndTable();

        // Set a uniform width (in points) for every cell in the table.
        double uniformWidth = 80.0; // points

        foreach (Row row in table.Rows)
        {
            foreach (Cell cell in row.Cells)
            {
                cell.CellFormat.Width = uniformWidth;
            }
        }

        // Save the document to the current directory.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "UniformColumnWidths.docx");
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("The output document was not saved correctly.");
    }
}
