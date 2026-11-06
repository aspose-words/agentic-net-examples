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

        // First cell.
        builder.InsertCell();
        builder.Writeln("English text (LTR)");

        // Second cell.
        builder.InsertCell();
        builder.Writeln("עברית (RTL)");

        // End the first row.
        builder.EndRow();

        // Second row (optional, just to have more cells).
        builder.InsertCell();
        builder.Writeln("More LTR text");
        builder.InsertCell();
        builder.Writeln("עוד טקסט RTL");
        builder.EndRow();

        // Finish the table.
        builder.EndTable();

        // Retrieve the created table.
        Table table = (Table)doc.GetChild(NodeType.Table, 0, true);

        // Set the text direction of every paragraph inside the table to right‑to‑left.
        foreach (Row row in table.Rows)
        {
            foreach (Cell cell in row.Cells)
            {
                foreach (Paragraph paragraph in cell.Paragraphs)
                {
                    paragraph.ParagraphFormat.Bidi = true; // Enable RTL layout.
                }
            }
        }

        // Save the document.
        string outputPath = "TableRtl.docx";
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("The output document was not saved correctly.");
    }
}
