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

        // Ensure the table has at least one row and one cell.
        table.EnsureMinimum();

        // Fill the table with sample text.
        for (int rowIdx = 0; rowIdx < 2; rowIdx++)
        {
            Row row = table.Rows[rowIdx];
            for (int cellIdx = 0; cellIdx < 2; cellIdx++)
            {
                Cell cell = row.Cells[cellIdx];
                cell.FirstParagraph.AppendChild(new Run(doc, $"R{rowIdx + 1}C{cellIdx + 1}"));
            }
        }

        // Apply read‑only protection to the entire document.
        // This makes the table (and all other content) non‑editable in Microsoft Word.
        doc.Protect(ProtectionType.ReadOnly);

        // Save the protected document.
        doc.Save("TableReadOnlyProtected.docx");
    }
}
