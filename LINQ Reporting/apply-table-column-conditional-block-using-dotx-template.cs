using System;
using Aspose.Words;
using Aspose.Words.Tables;
using Aspose.Words.Reporting;

class TableConditionalColumnExample
{
    static void Main()
    {
        // Load the DOTX template that contains a table.
        Document doc = new Document("Template.dotx");

        // Find the first table in the document.
        Table table = doc.FirstSection.Body.Tables[0];

        // Ensure the table has a style; if not, assign a built‑in style.
        if (table.Style == null)
            table.StyleIdentifier = StyleIdentifier.MediumShading1Accent1;

        // Access the table style's conditional styles collection.
        TableStyle tableStyle = (TableStyle)doc.Styles[table.StyleName];
        ConditionalStyleCollection condStyles = tableStyle.ConditionalStyles;

        // Apply a custom formatting to the odd column banding (e.g., light gray shading).
        condStyles.OddColumnBanding.Shading.BackgroundPatternColor = System.Drawing.Color.LightGray;

        // Enable the column banding option so the conditional style takes effect.
        table.StyleOptions |= TableStyleOptions.ColumnBands;

        // Optionally, enable other style options (first row, first column, etc.) as needed.
        // table.StyleOptions |= TableStyleOptions.FirstRow | TableStyleOptions.FirstColumn;

        // Save the modified document.
        doc.Save("Result.docx");
    }
}
