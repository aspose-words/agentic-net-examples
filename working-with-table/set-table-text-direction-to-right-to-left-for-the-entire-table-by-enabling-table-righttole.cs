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

        // Build a simple 2x1 table.
        builder.StartTable();
        builder.InsertCell();
        builder.Writeln("Cell 1");
        builder.InsertCell();
        builder.Writeln("Cell 2");
        builder.EndRow();
        builder.EndTable();

        // Retrieve the created table.
        Table table = (Table)doc.GetChild(NodeType.Table, 0, true);
        if (table == null)
            throw new InvalidOperationException("Table was not created.");

        // Enable right‑to‑left text direction for the entire table.
        // Aspose.Words does not expose a Table.RightToLeft property in this version,
        // so we set the Bidi flag on every paragraph inside each cell.
        foreach (Row row in table.Rows)
        {
            foreach (Cell cell in row.Cells)
            {
                foreach (Paragraph paragraph in cell.Paragraphs)
                {
                    paragraph.ParagraphFormat.Bidi = true;
                }
            }
        }

        // Save the document.
        string outputPath = "TableRightToLeft.docx";
        doc.Save(outputPath);

        // Verify that the file was saved.
        if (!File.Exists(outputPath))
            throw new FileNotFoundException("The output document was not saved.", outputPath);
    }
}
