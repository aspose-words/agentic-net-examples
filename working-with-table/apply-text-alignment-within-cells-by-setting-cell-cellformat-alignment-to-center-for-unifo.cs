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

        // First row, first cell.
        builder.InsertCell();
        builder.Writeln("Cell 1");
        // Center align text inside the cell by aligning the paragraph.
        builder.CurrentParagraph.ParagraphFormat.Alignment = ParagraphAlignment.Center;

        // First row, second cell.
        builder.InsertCell();
        builder.Writeln("Cell 2");
        builder.CurrentParagraph.ParagraphFormat.Alignment = ParagraphAlignment.Center;

        // End the first row.
        builder.EndRow();

        // Second row, first cell.
        builder.InsertCell();
        builder.Writeln("Cell 3");
        builder.CurrentParagraph.ParagraphFormat.Alignment = ParagraphAlignment.Center;

        // Second row, second cell.
        builder.InsertCell();
        builder.Writeln("Cell 4");
        builder.CurrentParagraph.ParagraphFormat.Alignment = ParagraphAlignment.Center;

        // End the second row and the table.
        builder.EndRow();
        builder.EndTable();

        // Save the document to disk.
        string outputPath = "AlignedTable.docx";
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
        {
            throw new InvalidOperationException($"Failed to create the output file: {outputPath}");
        }
    }
}
