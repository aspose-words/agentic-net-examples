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

        // Build a simple table with one cell containing Arabic text.
        builder.StartTable();
        builder.InsertCell();
        builder.Write("مرحبا بالعالم"); // Arabic greeting.
        builder.EndRow();
        builder.EndTable();

        // Retrieve the created table.
        Table table = (Table)doc.GetChild(NodeType.Table, 0, true);
        if (table == null)
            throw new InvalidOperationException("Table was not created.");

        // Enable right‑to‑left direction for the table by setting Bidi on each paragraph.
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
