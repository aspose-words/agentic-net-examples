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

        // Start building the table.
        builder.StartTable();

        // ----- Header row (bold) -----
        builder.Font.Bold = true;
        builder.InsertCell();
        builder.Writeln("Header 1");
        builder.InsertCell();
        builder.Writeln("Header 2");
        builder.EndRow();

        // ----- Data rows (regular) -----
        builder.Font.Bold = false;
        for (int i = 1; i <= 2; i++)
        {
            builder.InsertCell();
            builder.Writeln($"Data {i}A");
            builder.InsertCell();
            builder.Writeln($"Data {i}B");
            builder.EndRow();
        }

        // ----- Footer row (italic) -----
        builder.Font.Italic = true;
        builder.InsertCell();
        builder.Writeln("Footer 1");
        builder.InsertCell();
        builder.Writeln("Footer 2");
        builder.EndRow();

        // End the table.
        builder.EndTable();

        // Retrieve the created table.
        Table table = (Table)doc.GetChild(NodeType.Table, 0, true);

        // Apply a built‑in table style.
        table.StyleIdentifier = StyleIdentifier.TableGrid;

        // Save the document.
        string outputPath = "TableStyleExample.docx";
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("The output document was not created.");

        // The program finishes without waiting for user input.
    }
}
