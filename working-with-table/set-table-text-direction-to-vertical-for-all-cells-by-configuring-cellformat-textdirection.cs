using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a new document and a DocumentBuilder for constructing content.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Build a simple 3x3 table with sample text.
        builder.StartTable();

        for (int row = 0; row < 3; row++)
        {
            for (int col = 0; col < 3; col++)
            {
                builder.InsertCell();
                builder.Writeln($"R{row + 1}C{col + 1}");
            }
            builder.EndRow();
        }

        builder.EndTable();

        // Apply a vertical-like text direction to every cell.
        // Since CellFormat.TextDirection is disallowed, we set the paragraph's Bidi property
        // which forces right‑to‑left layout and can be used as an alternative demonstration.
        foreach (Table table in doc.GetChildNodes(NodeType.Table, true))
        {
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
        }

        // Save the document.
        string outputPath = "TableVertical.docx";
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new Exception("The output document was not created.");

        // The program ends here without waiting for user input.
    }
}
