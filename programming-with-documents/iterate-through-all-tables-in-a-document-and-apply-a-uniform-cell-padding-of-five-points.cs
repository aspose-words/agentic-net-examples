using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a new document and add a sample table.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Build a 2x2 table.
        builder.StartTable();
        builder.InsertCell();
        builder.Write("Cell 1");
        builder.InsertCell();
        builder.Write("Cell 2");
        builder.EndRow();
        builder.InsertCell();
        builder.Write("Cell 3");
        builder.InsertCell();
        builder.Write("Cell 4");
        builder.EndRow();
        builder.EndTable();

        // Apply a uniform cell padding of five points to every cell in every table.
        foreach (Table tbl in doc.GetChildNodes(NodeType.Table, true).OfType<Table>())
        {
            foreach (Row row in tbl.Rows)
            {
                foreach (Cell cell in row.Cells)
                {
                    cell.CellFormat.TopPadding = 5.0;
                    cell.CellFormat.BottomPadding = 5.0;
                    cell.CellFormat.LeftPadding = 5.0;
                    cell.CellFormat.RightPadding = 5.0;
                }
            }
        }

        // Save the document.
        string outputPath = "Output.docx";
        doc.Save(outputPath);

        // Simple verification that the file was created.
        if (File.Exists(outputPath))
        {
            Console.WriteLine($"Document saved successfully to {Path.GetFullPath(outputPath)}");
        }
        else
        {
            Console.WriteLine("Failed to save the document.");
        }
    }
}
