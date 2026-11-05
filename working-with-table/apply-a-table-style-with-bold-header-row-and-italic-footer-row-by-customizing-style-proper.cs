using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

public class TableStyleExample
{
    public static void Main()
    {
        // Create a new document and a DocumentBuilder for constructing content.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Define a custom table style.
        Style customStyle = doc.Styles.Add(StyleType.Table, "CustomTableStyle");
        // Set a base style (optional) – here we start from the built‑in "Table Grid".
        customStyle.BaseStyleName = "Table Grid";

        // Build a table with a header row, data rows, and a footer row.
        builder.StartTable();

        // ----- Header Row (bold) -----
        builder.Font.Bold = true; // Make header text bold.
        builder.InsertCell();
        builder.Writeln("Header 1");
        builder.InsertCell();
        builder.Writeln("Header 2");
        builder.EndRow();

        // ----- Data Row -----
        builder.Font.Bold = false; // Normal text for data rows.
        builder.InsertCell();
        builder.Writeln("Data 1");
        builder.InsertCell();
        builder.Writeln("Data 2");
        builder.EndRow();

        // ----- Footer Row (italic) -----
        builder.Font.Italic = true; // Make footer text italic.
        builder.InsertCell();
        builder.Writeln("Footer 1");
        builder.InsertCell();
        builder.Writeln("Footer 2");
        builder.EndRow();

        // End the table construction.
        builder.EndTable();

        // Retrieve the created table (the first table in the document).
        Table table = (Table)doc.GetChild(NodeType.Table, 0, true);

        // Apply the custom style to the table.
        table.Style = customStyle;

        // Save the document to a file.
        string outputPath = "TableStyleExample.docx";
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
        {
            throw new Exception("The output document was not created.");
        }
    }
}
