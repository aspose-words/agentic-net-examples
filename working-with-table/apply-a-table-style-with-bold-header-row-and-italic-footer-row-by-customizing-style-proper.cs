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

        // Build a simple 3‑row table (header, data, footer).
        builder.StartTable();

        // Header row.
        builder.InsertCell();
        builder.Write("Header 1");
        builder.InsertCell();
        builder.Write("Header 2");
        builder.EndRow();

        // Data row.
        builder.InsertCell();
        builder.Write("Data 1");
        builder.InsertCell();
        builder.Write("Data 2");
        builder.EndRow();

        // Footer row.
        builder.InsertCell();
        builder.Write("Footer 1");
        builder.InsertCell();
        builder.Write("Footer 2");
        builder.EndRow();

        // End the table.
        builder.EndTable();

        // Retrieve the created table.
        Table table = (Table)doc.GetChild(NodeType.Table, 0, true);

        // Create a custom table style.
        Style customStyle = doc.Styles.Add(StyleType.Table, "CustomTableStyle");
        // Set a base font for the whole table.
        customStyle.Font.Name = "Arial";

        // Apply the custom style to the table.
        // For a user‑defined style we only need to set the StyleName.
        table.StyleName = "CustomTableStyle";

        // Make header row bold.
        Row headerRow = table.FirstRow;
        foreach (Cell cell in headerRow.Cells)
        {
            Paragraph para = cell.FirstParagraph;
            Run run = para?.FirstChild as Run;
            if (run != null)
                run.Font.Bold = true;
        }

        // Make footer row italic.
        Row footerRow = table.LastRow;
        foreach (Cell cell in footerRow.Cells)
        {
            Paragraph para = cell.FirstParagraph;
            Run run = para?.FirstChild as Run;
            if (run != null)
                run.Font.Italic = true;
        }

        // Save the document.
        string outputPath = "TableStyleExample.docx";
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new Exception("The output document was not created.");
    }
}
