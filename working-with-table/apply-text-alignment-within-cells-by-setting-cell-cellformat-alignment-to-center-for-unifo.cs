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

        // Start a 2x2 table.
        Table table = builder.StartTable();

        // First row, first cell.
        builder.InsertCell();
        builder.Write("Cell 1, Row 1");
        // First row, second cell.
        builder.InsertCell();
        builder.Write("Cell 2, Row 1");
        builder.EndRow();

        // Second row, first cell.
        builder.InsertCell();
        builder.Write("Cell 1, Row 2");
        // Second row, second cell.
        builder.InsertCell();
        builder.Write("Cell 2, Row 2");
        builder.EndTable();

        // Apply horizontal center alignment to the paragraph inside each cell.
        foreach (Row row in table.Rows)
        {
            foreach (Cell cell in row.Cells)
            {
                // Ensure the cell contains at least one paragraph.
                cell.EnsureMinimum();

                // Set the alignment of the first paragraph to center.
                cell.FirstParagraph.ParagraphFormat.Alignment = ParagraphAlignment.Center;
            }
        }

        // Define output file path.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "AlignedTable.docx");

        // Save the document.
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
        {
            throw new Exception("The output file was not created.");
        }
    }
}
