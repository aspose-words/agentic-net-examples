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

        // First row (header).
        builder.InsertCell();
        builder.Write("Header 1");
        builder.InsertCell();
        builder.Write("Header 2");
        builder.EndRow();

        // Second row (data).
        builder.InsertCell();
        builder.Write("Row1 Col1");
        builder.InsertCell();
        builder.Write("Row1 Col2");
        builder.EndRow();

        // Finish the table.
        builder.EndTable();

        // Create a custom table style.
        TableStyle customStyle = (TableStyle)doc.Styles.Add(StyleType.Table, "MyTableStyle");

        // Make the first column bold via the conditional style.
        customStyle.ConditionalStyles[ConditionalStyleType.FirstColumn].Font.Bold = true;

        // Apply the custom style to the table.
        table.Style = customStyle;

        // Enable the first‑column conditional formatting.
        table.StyleOptions = TableStyleOptions.FirstColumn;

        // Save the document.
        string outputPath = "TableStyleFirstColumnBold.docx";
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new Exception("The output document was not saved correctly.");
    }
}
