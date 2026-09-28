using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a new document and a builder.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Build a simple table with one cell.
        builder.StartTable();
        builder.InsertCell();
        builder.Writeln("مرحبا بالعالم"); // Arabic text.
        builder.EndRow();
        builder.EndTable();

        // Locate the created cell.
        Table table = (Table)doc.GetChild(NodeType.Table, 0, true);
        Cell cell = table.Rows[0].Cells[0];

        // Set the paragraph direction inside the cell to right‑to‑left.
        foreach (Paragraph para in cell.Paragraphs)
        {
            para.ParagraphFormat.Bidi = true;
        }

        // Save the document.
        string outputPath = "CellDirection.docx";
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new Exception("The output file was not created.");
    }
}
