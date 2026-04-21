using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a new document and a builder to construct the table.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Define the size of the table.
        int rowCount = 3;
        int columnCount = 3;

        // Build the table with sample text in each cell.
        for (int r = 0; r < rowCount; r++)
        {
            for (int c = 0; c < columnCount; c++)
            {
                builder.InsertCell();
                builder.Write($"R{r + 1}C{c + 1}");
            }
            builder.EndRow();
        }

        // Finish the table and obtain a reference to it.
        Table table = builder.EndTable();

        // Desired uniform width for every column (in points).
        double uniformWidth = 100.0;

        // Iterate through each column and set the Width property of every cell in that column.
        for (int col = 0; col < columnCount; col++)
        {
            foreach (Row row in table.Rows)
            {
                Cell cell = row.Cells[col];
                cell.CellFormat.Width = uniformWidth;
            }
        }

        // Save the document to disk.
        string outputPath = "UniformColumnWidths.docx";
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new Exception("The document was not saved correctly.");
    }
}
