using System;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Tables;
using Aspose.Words.Settings;

class Program
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Set the document's compatibility level to Word 2016.
        doc.CompatibilityOptions.OptimizeFor(MsWordVersion.Word2016);

        // Begin a complex table.
        Table table = builder.StartTable();

        // ----- First row: merge the first three cells -----
        // Cell 1 (first of the merged range)
        builder.InsertCell();
        builder.Write("Header 1‑3 (merged)");
        builder.CellFormat.HorizontalMerge = CellMerge.First;
        // Cell 2 (previously merged)
        builder.InsertCell();
        builder.CellFormat.HorizontalMerge = CellMerge.Previous;
        // Cell 3 (previously merged)
        builder.InsertCell();
        builder.CellFormat.HorizontalMerge = CellMerge.Previous;
        // Remaining two header cells.
        builder.InsertCell();
        builder.Write("Header 4");
        builder.InsertCell();
        builder.Write("Header 5");
        builder.EndRow();

        // ----- Second row: regular cells -----
        for (int i = 0; i < 5; i++)
        {
            builder.InsertCell();
            builder.Write($"R2C{i + 1}");
        }
        builder.EndRow();

        // ----- Third row: merge the last two cells -----
        for (int i = 0; i < 3; i++)
        {
            builder.InsertCell();
            builder.Write($"R3C{i + 1}");
        }
        // Cell 4 – start of merged range
        builder.InsertCell();
        builder.Write("R3C4‑5 (merged)");
        builder.CellFormat.HorizontalMerge = CellMerge.First;
        // Cell 5 – continuation of merged range
        builder.InsertCell();
        builder.CellFormat.HorizontalMerge = CellMerge.Previous;
        builder.EndRow();

        // ----- Fourth row: apply shading to all cells -----
        for (int i = 0; i < 5; i++)
        {
            builder.InsertCell();
            builder.CellFormat.Shading.BackgroundPatternColor = Color.LightGray;
            builder.Write($"R4C{i + 1}");
        }
        builder.EndRow();

        // End the table.
        builder.EndTable();

        // Save the document.
        doc.Save("ComplexTable.docx");
    }
}
