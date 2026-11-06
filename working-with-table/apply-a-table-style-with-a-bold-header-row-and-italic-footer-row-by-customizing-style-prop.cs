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

        // Start a new table.
        builder.StartTable();

        // ----- Header Row -----
        builder.InsertCell();
        builder.Write("Header 1");
        builder.InsertCell();
        builder.Write("Header 2");
        builder.InsertCell();
        builder.Write("Header 3");
        builder.EndRow();

        // ----- Data Rows -----
        for (int i = 1; i <= 3; i++)
        {
            builder.InsertCell();
            builder.Write($"Data {i}A");
            builder.InsertCell();
            builder.Write($"Data {i}B");
            builder.InsertCell();
            builder.Write($"Data {i}C");
            builder.EndRow();
        }

        // ----- Footer Row -----
        builder.InsertCell();
        builder.Write("Footer 1");
        builder.InsertCell();
        builder.Write("Footer 2");
        builder.InsertCell();
        builder.Write("Footer 3");
        builder.EndRow();

        // End the table.
        builder.EndTable();

        // Retrieve the created table (the first table in the document).
        Table table = (Table)doc.GetChild(NodeType.Table, 0, true);

        // Create a custom table style.
        Style customStyle = doc.Styles.Add(StyleType.Table, "CustomTableStyle");

        // Apply the style to the table.
        table.Style = customStyle;

        // Make the first row (header) bold via style options.
        // The TableStyleOptions property belongs to Table, not Style.
        table.StyleOptions = TableStyleOptions.FirstRow;
        customStyle.Font.Bold = true;

        // Make the last row (footer) italic by formatting its runs directly.
        Row footerRow = table.Rows[table.Rows.Count - 1];
        foreach (Cell cell in footerRow.Cells)
        {
            foreach (Paragraph para in cell.Paragraphs)
            {
                foreach (Run run in para.Runs)
                {
                    run.Font.Italic = true;
                }
            }
        }

        // Save the document.
        string outputPath = "TableStyleExample.docx";
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("The output document was not created.");
    }
}
