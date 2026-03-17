using Aspose.Words;
using Aspose.Words.Tables;
using System;

class Program
{
    static void Main()
    {
        // Create a new document (lifecycle: create)
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Build a simple 3‑row, 2‑column table
        Table table = builder.StartTable();
        for (int r = 0; r < 3; r++)
        {
            for (int c = 0; c < 2; c++)
            {
                builder.InsertCell();
                builder.Writeln($"Row {r + 1}, Cell {c + 1}");
            }
            builder.EndRow(); // end current row
        }
        builder.EndTable();

        // Adjust spacing before and after each row.
        // RowFormat does not expose SpaceBefore/SpaceAfter, so we set the paragraph
        // spacing of the first paragraph in each cell, which effectively adds space
        // above and below the row content.
        foreach (Row row in table.Rows)
        {
            foreach (Cell cell in row.Cells)
            {
                Paragraph para = cell.FirstParagraph;
                para.ParagraphFormat.SpaceBefore = 12; // points before the row
                para.ParagraphFormat.SpaceAfter = 12;  // points after the row
            }
        }

        // Save the document (lifecycle: save)
        doc.Save("RowSpacing.docx");
    }
}
