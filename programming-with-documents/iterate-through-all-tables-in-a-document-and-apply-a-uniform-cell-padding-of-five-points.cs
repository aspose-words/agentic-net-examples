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

        // Build a sample table with two rows and two columns.
        Table table = builder.StartTable();
        builder.InsertCell();
        builder.Write("Cell 1");
        builder.InsertCell();
        builder.Write("Cell 2");
        builder.EndRow();

        builder.InsertCell();
        builder.Write("Cell 3");
        builder.InsertCell();
        builder.Write("Cell 4");
        builder.EndTable();

        // Iterate through all tables in the document.
        NodeCollection tables = doc.GetChildNodes(NodeType.Table, true);
        foreach (Table tbl in tables)
        {
            // Iterate through each cell of the current table and set uniform padding of 5 points.
            foreach (Row row in tbl.Rows)
            {
                foreach (Cell cell in row.Cells)
                {
                    cell.CellFormat.SetPaddings(5, 5, 5, 5);
                }
            }
        }

        // Define an output path and ensure the directory exists.
        string outputDir = Path.Combine(Environment.CurrentDirectory, "Output");
        Directory.CreateDirectory(outputDir);
        string outputPath = Path.Combine(outputDir, "PaddedTables.docx");

        // Save the modified document.
        doc.Save(outputPath);
    }
}
