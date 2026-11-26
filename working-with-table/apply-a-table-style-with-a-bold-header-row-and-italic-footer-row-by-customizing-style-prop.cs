using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a new document and a DocumentBuilder.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Start the table.
        builder.StartTable();

        // Header row.
        builder.InsertCell();
        builder.Writeln("Header 1");
        builder.InsertCell();
        builder.Writeln("Header 2");
        builder.EndRow();

        // Data row 1.
        builder.InsertCell();
        builder.Writeln("Data 1A");
        builder.InsertCell();
        builder.Writeln("Data 1B");
        builder.EndRow();

        // Data row 2.
        builder.InsertCell();
        builder.Writeln("Data 2A");
        builder.InsertCell();
        builder.Writeln("Data 2B");
        builder.EndRow();

        // Footer row.
        builder.InsertCell();
        builder.Writeln("Footer 1");
        builder.InsertCell();
        builder.Writeln("Footer 2");
        builder.EndRow();

        // End the table.
        builder.EndTable();

        // Retrieve the created table.
        Table table = (Table)doc.GetChildNodes(NodeType.Table, true)[0];

        // Apply bold formatting to the header row.
        Row headerRow = table.FirstRow;
        foreach (Cell cell in headerRow.Cells)
        {
            // Ensure the cell contains at least one paragraph and one run.
            if (cell.FirstParagraph?.Runs.Count > 0)
            {
                cell.FirstParagraph.Runs[0].Font.Bold = true;
            }
        }

        // Apply italic formatting to the footer row.
        Row footerRow = table.LastRow;
        foreach (Cell cell in footerRow.Cells)
        {
            if (cell.FirstParagraph?.Runs.Count > 0)
            {
                cell.FirstParagraph.Runs[0].Font.Italic = true;
            }
        }

        // Optionally, assign a simple table style (e.g., "Table Grid") to keep a base style.
        table.StyleIdentifier = StyleIdentifier.TableGrid;

        // Save the document.
        string outputPath = "TableStyleExample.docx";
        doc.Save(outputPath);

        // Verify that the file was saved.
        if (!File.Exists(outputPath))
        {
            throw new Exception($"Failed to create the output file: {outputPath}");
        }

        Console.WriteLine($"Document saved successfully to '{outputPath}'.");
    }
}
