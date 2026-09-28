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
        Table table = builder.StartTable();

        // ----- Header row -----
        builder.InsertCell();
        builder.Writeln("Header 1");
        builder.InsertCell();
        builder.Writeln("Header 2");
        builder.EndRow();

        // ----- Data rows -----
        for (int i = 1; i <= 3; i++)
        {
            builder.InsertCell();
            builder.Writeln($"Data {i}A");
            builder.InsertCell();
            builder.Writeln($"Data {i}B");
            builder.EndRow();
        }

        // ----- Footer row -----
        builder.InsertCell();
        builder.Writeln("Footer 1");
        builder.InsertCell();
        builder.Writeln("Footer 2");
        builder.EndRow();

        // End the table.
        builder.EndTable();

        // Create a custom table style based on the built‑in "Table Grid" style.
        Style customStyle = doc.Styles.Add(StyleType.Table, "MyCustomTableStyle");
        customStyle.BaseStyleName = "Table Grid";

        // Apply the custom style to the table.
        // Setting StyleName is sufficient; no need to set StyleIdentifier.
        table.StyleName = "MyCustomTableStyle";

        // Make header row text bold.
        Row headerRow = table.Rows[0];
        foreach (Cell cell in headerRow.Cells)
        {
            Paragraph para = cell.FirstParagraph;
            if (para != null && para.Runs.Count > 0)
                para.Runs[0].Font.Bold = true;
        }

        // Make footer row text italic.
        Row footerRow = table.Rows[table.Rows.Count - 1];
        foreach (Cell cell in footerRow.Cells)
        {
            Paragraph para = cell.FirstParagraph;
            if (para != null && para.Runs.Count > 0)
                para.Runs[0].Font.Italic = true;
        }

        // Save the document.
        string outputPath = "TableStyleExample.docx";
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new Exception("The output document was not created.");
    }
}
