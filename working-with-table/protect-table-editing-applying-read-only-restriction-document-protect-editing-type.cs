using System;
using Aspose.Words;
using Aspose.Words.Tables;

class TableReadOnlyProtection
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Build a simple 2x2 table.
        Table table = new Table(doc);
        doc.FirstSection.Body.AppendChild(table);

        // Add rows and cells with sample text.
        for (int rowIdx = 0; rowIdx < 2; rowIdx++)
        {
            Row row = new Row(doc);
            table.AppendChild(row);

            for (int cellIdx = 0; cellIdx < 2; cellIdx++)
            {
                Cell cell = new Cell(doc);
                // Each cell must contain at least one paragraph.
                cell.AppendChild(new Paragraph(doc));
                row.AppendChild(cell);

                cell.FirstParagraph.AppendChild(
                    new Run(doc, $"R{rowIdx + 1}C{cellIdx + 1}"));
            }
        }

        // Apply read‑only protection to the entire document.
        // This makes the table (and all other content) non‑editable in Microsoft Word.
        doc.Protect(ProtectionType.ReadOnly);

        // Save the protected document.
        doc.Save("TableReadOnlyProtected.docx");
    }
}
