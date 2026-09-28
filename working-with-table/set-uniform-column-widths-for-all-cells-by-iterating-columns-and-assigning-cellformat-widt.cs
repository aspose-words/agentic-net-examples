using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a new document and a DocumentBuilder.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Build a 3x3 table.
        builder.StartTable();
        for (int row = 0; row < 3; row++)
        {
            for (int col = 0; col < 3; col++)
            {
                builder.InsertCell();
                builder.Writeln($"R{row + 1}C{col + 1}");
            }
            builder.EndRow();
        }
        builder.EndTable();

        // Retrieve the created table.
        Table table = (Table)doc.GetChild(NodeType.Table, 0, true);

        // Desired uniform column width (points).
        double uniformWidth = 80.0;

        // Set the same width for each cell in every column.
        int columnCount = table.Rows[0].Cells.Count;
        for (int colIndex = 0; colIndex < columnCount; colIndex++)
        {
            foreach (Row row in table.Rows)
            {
                Cell cell = row.Cells[colIndex];
                cell.CellFormat.Width = uniformWidth;
            }
        }

        // Save the document.
        string outputPath = "UniformColumnWidths.docx";
        doc.Save(outputPath);

        // Verify that the file was saved.
        if (!File.Exists(outputPath))
            throw new Exception("Document was not saved successfully.");

        Console.WriteLine($"Document saved to {outputPath}");
    }
}
