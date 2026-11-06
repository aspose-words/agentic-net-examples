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

        // Build a simple 3‑row, 2‑column table.
        builder.StartTable();

        // Header row.
        builder.InsertCell();
        builder.Write("Header 1");
        builder.InsertCell();
        builder.Write("Header 2");
        builder.EndRow();

        // First data row.
        builder.InsertCell();
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

        // Finish the table.
        builder.EndTable();

        // Retrieve the created table.
        Table table = (Table)doc.GetChild(NodeType.Table, 0, true);
        if (table == null)
            throw new InvalidOperationException("Table was not created.");

        // Apply a built‑in style (e.g., "Table Grid").
        table.Style = doc.Styles["Table Grid"];

        // Enable the first column style option.
        table.StyleOptions = TableStyleOptions.FirstColumn;

        // Apply bold formatting to all cells in the first column.
        foreach (Row row in table.Rows)
        {
            Cell firstCell = row.Cells[0];
            foreach (Paragraph para in firstCell.Paragraphs)
            {
                foreach (Run run in para.Runs)
                {
                    run.Font.Bold = true;
                }
            }
        }

        // Save the document.
        string outputPath = "TableStyleFirstColumnBold.docx";
        doc.Save(outputPath);

        // Verify that the file was saved.
        if (!File.Exists(outputPath))
            throw new FileNotFoundException("The output document was not created.", outputPath);
    }
}
