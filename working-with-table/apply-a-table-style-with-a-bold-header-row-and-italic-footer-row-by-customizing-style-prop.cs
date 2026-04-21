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

        // Start building a table.
        Table table = builder.StartTable();

        // ----- Header row (will be styled bold) -----
        builder.InsertCell();
        builder.Write("Header 1");
        builder.InsertCell();
        builder.Write("Header 2");
        builder.EndRow();

        // ----- Data rows -----
        for (int i = 1; i <= 3; i++)
        {
            builder.InsertCell();
            builder.Write($"Item {i}");
            builder.InsertCell();
            builder.Write($"Value {i}");
            builder.EndRow();
        }

        // ----- Footer row (will be styled italic) -----
        builder.InsertCell();
        builder.Write("Footer 1");
        builder.InsertCell();
        builder.Write("Footer 2");
        builder.EndRow();

        // Finish the table.
        builder.EndTable();

        // Create a custom table style.
        TableStyle customStyle = (TableStyle)doc.Styles.Add(StyleType.Table, "MyCustomStyle");

        // Apply bold formatting to the first (header) row.
        customStyle.ConditionalStyles[ConditionalStyleType.FirstRow].Font.Bold = true;

        // Apply italic formatting to the last (footer) row.
        customStyle.ConditionalStyles[ConditionalStyleType.LastRow].Font.Italic = true;

        // Assign the style to the table.
        table.Style = customStyle;

        // Enable the conditional formatting for first and last rows.
        table.StyleOptions = TableStyleOptions.FirstRow | TableStyleOptions.LastRow;

        // Save the document to the local file system.
        string outputPath = "TableStyleHeaderFooter.docx";
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new Exception("The output document was not created.");
    }
}
