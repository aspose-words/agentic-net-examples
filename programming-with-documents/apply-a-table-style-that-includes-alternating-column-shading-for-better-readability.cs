using System;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Build a simple 4x4 table.
        builder.StartTable();
        for (int row = 0; row < 4; row++)
        {
            for (int col = 0; col < 4; col++)
            {
                builder.InsertCell();
                builder.Writeln($"R{row + 1}C{col + 1}");
            }
            builder.EndRow();
        }
        builder.EndTable();

        // Retrieve the created table (it is the node before the current paragraph).
        Table table = (Table)builder.CurrentParagraph.PreviousSibling;

        // Apply a built‑in style that supports column banding.
        table.StyleIdentifier = StyleIdentifier.TableGridLight;

        // Enable alternating column shading.
        table.StyleOptions = TableStyleOptions.ColumnBands;

        // Save the document to disk.
        string outputPath = "TableWithColumnBanding.docx";
        doc.Save(outputPath);
    }
}
