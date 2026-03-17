using System;
using Aspose.Words;
using Aspose.Words.Tables;

class Program
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Set the spacing before and after the table (points).
        // In Aspose.Words the spacing is a paragraph property, not a Table property.
        builder.ParagraphFormat.SpaceBefore = 12.0; // 12 points = 1/6 inch
        builder.ParagraphFormat.SpaceAfter = 12.0;

        // Build a simple 1‑row, 2‑cell table.
        Table table = builder.StartTable();
        builder.InsertCell();
        builder.Write("Cell 1");
        builder.InsertCell();
        builder.Write("Cell 2");
        builder.EndRow();
        builder.EndTable();

        // Save the document to disk.
        doc.Save("TableSpacing.docx");
    }
}
