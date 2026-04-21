using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

public class TableStyleFirstColumnBoldExample
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Build a simple 3x2 table.
        Table table = builder.StartTable();

        // First row (header).
        builder.InsertCell();
        builder.Write("Item");
        builder.InsertCell();
        builder.Write("Quantity");
        builder.EndRow();

        // First data row.
        builder.InsertCell();
        builder.Write("Apples");
        builder.InsertCell();
        builder.Write("10");
        builder.EndRow();

        // Second data row.
        builder.InsertCell();
        builder.Write("Bananas");
        builder.InsertCell();
        builder.Write("20");
        builder.EndRow();

        // Finish the table.
        builder.EndTable();

        // Create a custom table style.
        TableStyle customStyle = (TableStyle)doc.Styles.Add(StyleType.Table, "MyCustomStyle");
        // Make the first column bold via the conditional style.
        customStyle.ConditionalStyles[ConditionalStyleType.FirstColumn].Font.Bold = true;

        // Apply the custom style to the table.
        table.Style = customStyle;

        // Enable the FirstColumn style option so the conditional style is applied.
        table.StyleOptions = TableStyleOptions.FirstColumn;

        // Save the document.
        string outputPath = Path.Combine(Environment.CurrentDirectory, "TableStyleFirstColumnBold.docx");
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("The output document was not saved correctly.");

        // The program ends here without waiting for user input.
    }
}
