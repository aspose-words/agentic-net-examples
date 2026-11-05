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

        // Build a simple 2x3 table.
        builder.StartTable();

        // First row.
        for (int i = 0; i < 3; i++)
        {
            builder.InsertCell();
            builder.Writeln($"Cell {i + 1}");
        }
        builder.EndRow();

        // Second row.
        for (int i = 0; i < 3; i++)
        {
            builder.InsertCell();
            builder.Writeln($"Cell {i + 4}");
        }
        builder.EndRow();

        builder.EndTable();

        // Center-align the paragraphs inside each cell.
        foreach (Table table in doc.GetChildNodes(NodeType.Table, true))
        {
            foreach (Row row in table.Rows)
            {
                foreach (Cell cell in row.Cells)
                {
                    foreach (Paragraph para in cell.Paragraphs)
                    {
                        para.ParagraphFormat.Alignment = ParagraphAlignment.Center;
                    }
                }
            }
        }

        // Save the document.
        string outputPath = "AlignedTable.docx";
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new Exception("The output file was not created.");
    }
}
