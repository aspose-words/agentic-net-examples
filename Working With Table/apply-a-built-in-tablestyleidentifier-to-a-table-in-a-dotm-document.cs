using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

class Program
{
    static void Main()
    {
        // Load an existing DOTM template
        string dataDir = @"C:\Data\";
        string inputPath = Path.Combine(dataDir, "Template.dotm");
        Document doc = new Document(inputPath);

        // Use DocumentBuilder to insert a new table
        DocumentBuilder builder = new DocumentBuilder(doc);
        Table table = builder.StartTable();

        // Insert header row
        builder.InsertCell();
        builder.Writeln("Header 1");
        builder.InsertCell();
        builder.Writeln("Header 2");
        builder.EndRow();

        // Insert data row
        builder.InsertCell();
        builder.Writeln("Value 1");
        builder.InsertCell();
        builder.Writeln("Value 2");
        builder.EndRow();

        // Finish table construction
        builder.EndTable();

        // Apply a built‑in table style by its identifier
        table.StyleIdentifier = StyleIdentifier.MediumShading1Accent1;

        // Enable specific style features (first row, first column, row banding)
        table.StyleOptions = TableStyleOptions.FirstRow |
                             TableStyleOptions.FirstColumn |
                             TableStyleOptions.RowBands;

        // Adjust column widths to fit the content
        table.AutoFit(AutoFitBehavior.AutoFitToContents);

        // Save the modified document as a DOTM file
        string outputPath = Path.Combine(dataDir, "StyledTable.dotm");
        doc.Save(outputPath);
    }
}
