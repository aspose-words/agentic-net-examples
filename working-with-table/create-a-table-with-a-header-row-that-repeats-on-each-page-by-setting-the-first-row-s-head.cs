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

        // Start the table.
        builder.StartTable();

        // ---- Header row (will repeat on each page) ----
        builder.InsertCell();
        builder.Writeln("Header 1");
        builder.InsertCell();
        builder.Writeln("Header 2");
        builder.InsertCell();
        builder.Writeln("Header 3");
        builder.EndRow();

        // Retrieve the table that has just been created.
        // After ending the row the builder is positioned inside the last cell's paragraph,
        // so the hierarchy is: Paragraph -> Cell -> Row -> Table.
        Table table = (Table)builder.CurrentParagraph.ParentNode.ParentNode.ParentNode;

        // Mark the first row as a heading row so it repeats on each page.
        table.Rows[0].RowFormat.HeadingFormat = true;

        // ---- Add enough data rows to span multiple pages ----
        for (int i = 1; i <= 50; i++)
        {
            builder.InsertCell();
            builder.Writeln($"Row {i} Col 1");
            builder.InsertCell();
            builder.Writeln($"Row {i} Col 2");
            builder.InsertCell();
            builder.Writeln($"Row {i} Col 3");
            builder.EndRow();
        }

        // End the table.
        builder.EndTable();

        // Save the document.
        string outputPath = "TableWithRepeatingHeader.docx";
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new Exception("Failed to create the output document.");
    }
}
