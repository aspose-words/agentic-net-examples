using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Build a simple 1x1 table with Arabic text.
        builder.StartTable();
        builder.InsertCell();
        builder.Writeln("مرحبا بالعالم"); // Arabic greeting.
        builder.EndRow();
        builder.EndTable();

        // Retrieve the created table.
        Table table = (Table)doc.GetChild(NodeType.Table, 0, true);
        if (table == null)
            throw new InvalidOperationException("Table was not created.");

        // Enable right‑to‑left direction for the table by setting Bidi on all paragraphs.
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

        // Verify that the file was saved.
        if (!File.Exists(outputPath))
            throw new FileNotFoundException("The output file was not created.", outputPath);
    }
}
