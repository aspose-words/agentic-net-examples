using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a new document and a DocumentBuilder for building content.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Build a simple 2x2 table with sample text.
        builder.StartTable();
        builder.InsertCell();
        builder.Writeln("Cell 1,1");
        builder.InsertCell();
        builder.Writeln("Cell 1,2");
        builder.EndRow();

        builder.InsertCell();
        builder.Writeln("Cell 2,1");
        builder.InsertCell();
        builder.Writeln("Cell 2,2");
        builder.EndRow();
        builder.EndTable();

        // Remove all borders from every table in the document.
        foreach (Table table in doc.GetChildNodes(NodeType.Table, true))
        {
            // Remove borders from each row.
            foreach (Row row in table.Rows)
            {
                row.RowFormat.Borders.LineStyle = LineStyle.None;

                // Remove borders from each cell in the row.
                foreach (Cell cell in row.Cells)
                {
                    cell.CellFormat.Borders.LineStyle = LineStyle.None;
                }
            }
        }

        // Save the document to a file.
        string outputPath = "Output.docx";
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
        {
            throw new Exception("The output file was not created.");
        }
    }
}
