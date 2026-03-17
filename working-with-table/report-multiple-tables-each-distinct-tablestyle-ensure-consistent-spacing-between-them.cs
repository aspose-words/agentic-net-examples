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

        // Define distinct built‑in table styles to apply to each table.
        StyleIdentifier[] tableStyles = new StyleIdentifier[]
        {
            StyleIdentifier.TableGrid,               // Simple grid style.
            StyleIdentifier.LightListAccent1,        // Light list with accent 1.
            StyleIdentifier.MediumShading1Accent2    // Medium shading with accent 2.
        };

        // Insert each table, applying its own style, and add uniform spacing between them.
        for (int i = 0; i < tableStyles.Length; i++)
        {
            InsertStyledTable(builder, tableStyles[i]);

            // Insert an empty paragraph to create consistent spacing after the table.
            // The default paragraph spacing will be used for all tables.
            builder.Writeln();
        }

        // Save the resulting document.
        doc.Save("MultipleTablesWithStyles.docx");
    }

    // Creates a simple 2×2 table, applies the specified style, and sets common formatting.
    static void InsertStyledTable(DocumentBuilder builder, StyleIdentifier styleId)
    {
        // Begin a new table.
        Table table = builder.StartTable();

        // First row – header cells.
        builder.InsertCell();
        builder.Write("Header 1");
        builder.InsertCell();
        builder.Write("Header 2");
        builder.EndRow();

        // Second row – data cells.
        builder.InsertCell();
        builder.Write("Data 1");
        builder.InsertCell();
        builder.Write("Data 2");
        builder.EndRow();

        // Finish the table.
        builder.EndTable();

        // Apply the chosen built‑in style.
        table.StyleIdentifier = styleId;

        // Enable common style options (first row formatting and row banding).
        table.StyleOptions = TableStyleOptions.FirstRow | TableStyleOptions.RowBands;

        // Ensure consistent cell spacing for all tables.
        table.AllowCellSpacing = true;
        table.CellSpacing = 3.0; // 3 points between cells.

        // Auto‑fit the table to its contents for a compact layout.
        table.AutoFit(AutoFitBehavior.AutoFitToContents);
    }
}
