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

        // First table – Light Shading Accent 1 style.
        BuildSampleTable(builder, StyleIdentifier.LightShadingAccent1, "Table 1");

        // Add a blank paragraph to create consistent spacing between tables.
        builder.Writeln();

        // Second table – Medium Shading 1 Accent 2 style.
        BuildSampleTable(builder, StyleIdentifier.MediumShading1Accent2, "Table 2");

        // Add spacing.
        builder.Writeln();

        // Third table – Medium Shading 2 Accent 3 style.
        BuildSampleTable(builder, StyleIdentifier.MediumShading2Accent3, "Table 3");

        // Save the document.
        doc.Save("ReportWithMultipleTables.docx");
    }

    // Helper method that builds a simple 3‑row table and applies the specified style.
    private static void BuildSampleTable(DocumentBuilder builder, StyleIdentifier styleId, string title)
    {
        // Start the table.
        Table table = builder.StartTable();

        // Header row.
        builder.InsertCell();
        builder.Writeln($"{title} Header 1");
        builder.InsertCell();
        builder.Writeln($"{title} Header 2");
        builder.EndRow();

        // First data row.
        builder.InsertCell();
        builder.Writeln($"{title} Row 1, Col 1");
        builder.InsertCell();
        builder.Writeln($"{title} Row 1, Col 2");
        builder.EndRow();

        // Second data row.
        builder.InsertCell();
        builder.Writeln($"{title} Row 2, Col 1");
        builder.InsertCell();
        builder.Writeln($"{title} Row 2, Col 2");
        builder.EndRow();

        // Apply the requested style to the table.
        table.StyleIdentifier = styleId;
        // Apply first‑row formatting and row banding for visual distinction.
        table.StyleOptions = TableStyleOptions.FirstRow | TableStyleOptions.RowBands;
        // Resize the table to fit its contents.
        table.AutoFit(AutoFitBehavior.AutoFitToContents);

        // End the table.
        builder.EndTable();
    }
}
