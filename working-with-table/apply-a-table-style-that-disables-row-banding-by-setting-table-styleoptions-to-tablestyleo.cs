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

        // First row (header).
        builder.InsertCell();
        builder.Write("Header 1");
        builder.InsertCell();
        builder.Write("Header 2");
        builder.EndRow();

        // Second row (data).
        builder.InsertCell();
        builder.Write("Data 1");
        builder.InsertCell();
        builder.Write("Data 2");
        builder.EndTable();

        // Apply a built‑in style (optional) and disable row banding.
        table.StyleIdentifier = StyleIdentifier.LightShadingAccent1;
        // Setting StyleOptions to None removes all conditional formatting, including row banding.
        table.StyleOptions = TableStyleOptions.None;

        // Save the document.
        string outputPath = Path.Combine(Environment.CurrentDirectory, "Table_NoRowBanding.docx");
        doc.Save(outputPath);
    }
}
