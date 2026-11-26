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

        // Build a 2x2 table.
        builder.StartTable();

        // First row.
        builder.InsertCell();
        builder.Writeln("Cell 1,1");
        builder.InsertCell();
        builder.Writeln("Cell 1,2");
        builder.EndRow();

        // Second row.
        builder.InsertCell();
        builder.Writeln("Cell 2,1");
        builder.InsertCell();
        builder.Writeln("Cell 2,2");
        builder.EndRow();

        // End the table.
        builder.EndTable();

        // Retrieve the created table (the first table in the document).
        Table table = (Table)doc.GetChild(NodeType.Table, 0, true);

        // Ensure all paragraphs inside the table are marked as right‑to‑left.
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
        string outputPath = "TableRtl.docx";
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new Exception("The output file was not created.");

        // Load the document again to confirm the setting persists.
        Document loadedDoc = new Document(outputPath);
        Table loadedTable = (Table)loadedDoc.GetChild(NodeType.Table, 0, true);

        // Verify that every paragraph inside the loaded table has Bidi set.
        foreach (Row row in loadedTable.Rows)
        {
            foreach (Cell cell in row.Cells)
            {
                foreach (Paragraph para in cell.Paragraphs)
                {
                    if (!para.ParagraphFormat.Bidi)
                        throw new Exception("A paragraph inside the table is not set to right‑to‑left.");
                }
            }
        }

        // If we reach this point, the operation succeeded.
        Console.WriteLine("Table text direction set to right‑to‑left and verified successfully.");
    }
}
