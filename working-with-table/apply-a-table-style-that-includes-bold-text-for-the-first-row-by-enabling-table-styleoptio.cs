using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Build a simple 2x2 table.
        Table table = builder.StartTable();

        // First row – this row will be styled to appear bold.
        builder.InsertCell();
        builder.Write("Header 1");
        builder.InsertCell();
        builder.Write("Header 2");
        builder.EndRow();

        // Second row – regular formatting.
        builder.InsertCell();
        builder.Write("Value 1");
        builder.InsertCell();
        builder.Write("Value 2");
        builder.EndRow();

        // Finish the table.
        builder.EndTable();

        // Apply a built‑in table style.
        table.StyleIdentifier = StyleIdentifier.LightShadingAccent1;

        // Enable the conditional formatting for the first row.
        table.StyleOptions = TableStyleOptions.FirstRow;

        // Retrieve the style object that was applied to the table.
        TableStyle appliedStyle = (TableStyle)doc.Styles[table.StyleIdentifier];

        // Make the text in the first row bold via the conditional style.
        appliedStyle.ConditionalStyles[ConditionalStyleType.FirstRow].Font.Bold = true;

        // Save the document.
        string outputPath = "TableStyleFirstRowBold.docx";
        doc.Save(outputPath);

        // Simple validation to ensure the file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("The output document was not saved correctly.");
    }
}
