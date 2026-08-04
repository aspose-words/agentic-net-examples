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

        // ---------- Header row ----------
        builder.InsertCell();
        builder.Write("Header 1");
        builder.InsertCell();
        builder.Write("Header 2");
        builder.EndRow();

        // ---------- Data rows ----------
        for (int i = 1; i <= 2; i++)
        {
            builder.InsertCell();
            builder.Write($"Data {i}A");
            builder.InsertCell();
            builder.Write($"Data {i}B");
            builder.EndRow();
        }

        // ---------- Footer row ----------
        builder.InsertCell();
        builder.Write("Footer 1");
        builder.InsertCell();
        builder.Write("Footer 2");
        builder.EndRow();

        // Finish the table.
        builder.EndTable();

        // Create a custom table style.
        TableStyle customStyle = (TableStyle)doc.Styles.Add(StyleType.Table, "MyTableStyle");

        // Make the first row (header) bold.
        customStyle.ConditionalStyles[ConditionalStyleType.FirstRow].Font.Bold = true;

        // Make the last row (footer) italic.
        customStyle.ConditionalStyles[ConditionalStyleType.LastRow].Font.Italic = true;

        // Apply the style to the table and enable the conditional formatting for header and footer.
        table.Style = customStyle;
        table.StyleOptions = TableStyleOptions.FirstRow | TableStyleOptions.LastRow;

        // Save the document.
        string outputPath = Path.Combine(Environment.CurrentDirectory, "TableStyleHeaderFooter.docx");
        doc.Save(outputPath);

        // Simple verification that the file was created.
        if (!File.Exists(outputPath))
            throw new Exception("The output document was not saved correctly.");
    }
}
