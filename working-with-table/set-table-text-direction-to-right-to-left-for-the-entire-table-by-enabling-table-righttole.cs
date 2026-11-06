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

        // Build a simple 2x2 table.
        builder.StartTable();

        // First row, first cell.
        builder.InsertCell();
        builder.Write("Cell 1");

        // First row, second cell.
        builder.InsertCell();
        builder.Write("Cell 2");
        builder.EndRow();

        // Second row, first cell.
        builder.InsertCell();
        builder.Write("Cell 3");

        // Second row, second cell.
        builder.InsertCell();
        builder.Write("Cell 4");
        builder.EndRow();

        // End the table construction.
        builder.EndTable();

        // Retrieve the created table.
        Table table = (Table)doc.GetChildNodes(NodeType.Table, true)[0];

        // Set right‑to‑left text direction for the entire table by enabling Bidi on all paragraphs.
        foreach (Row row in table.Rows)
        {
            foreach (Cell cell in row.Cells)
            {
                foreach (Paragraph para in cell.Paragraphs)
                {
                    para.ParagraphFormat.Bidi = true;
                }
            }
        }

        // Save the document.
        string outputPath = "TableRightToLeft.docx";
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new Exception("The output document was not saved correctly.");
    }
}
