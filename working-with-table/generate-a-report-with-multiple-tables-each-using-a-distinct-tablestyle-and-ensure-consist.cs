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

        // Apply a consistent space after each table (12 points).
        builder.ParagraphFormat.SpaceAfter = 12;

        // ---------- First Table ----------
        Table table1 = builder.StartTable();

        // Insert a cell to ensure the table has at least one row before applying style.
        builder.InsertCell();
        table1.StyleIdentifier = StyleIdentifier.LightShadingAccent1;
        table1.StyleOptions = TableStyleOptions.FirstRow | TableStyleOptions.RowBands;

        // Header row
        builder.Writeln("Header 1");
        builder.InsertCell();
        builder.Writeln("Header 2");
        builder.EndRow();

        // Data rows
        builder.InsertCell();
        builder.Writeln("Row1 Col1");
        builder.InsertCell();
        builder.Writeln("Row1 Col2");
        builder.EndRow();

        builder.InsertCell();
        builder.Writeln("Row2 Col1");
        builder.InsertCell();
        builder.Writeln("Row2 Col2");
        builder.EndRow();

        builder.EndTable();

        // Add a blank paragraph to separate tables.
        builder.Writeln();

        // ---------- Second Table ----------
        Table table2 = builder.StartTable();
        builder.InsertCell(); // ensure at least one row before styling
        table2.StyleIdentifier = StyleIdentifier.MediumShading1Accent1;
        table2.StyleOptions = TableStyleOptions.FirstRow | TableStyleOptions.RowBands;

        // Header row
        builder.Writeln("Product");
        builder.InsertCell();
        builder.Writeln("Qty");
        builder.InsertCell();
        builder.Writeln("Price");
        builder.EndRow();

        // Sample data rows
        for (int i = 1; i <= 2; i++)
        {
            builder.InsertCell();
            builder.Writeln($"Item {i}");
            builder.InsertCell();
            builder.Writeln((i * 10).ToString());
            builder.InsertCell();
            builder.Writeln($"${i * 5}.00");
            builder.EndRow();
        }

        builder.EndTable();

        builder.Writeln();

        // ---------- Third Table ----------
        Table table3 = builder.StartTable();
        builder.InsertCell(); // ensure at least one row before styling
        table3.StyleIdentifier = StyleIdentifier.DarkList;
        table3.StyleOptions = TableStyleOptions.FirstRow | TableStyleOptions.RowBands;

        // Single row with four cells
        builder.Writeln("A");
        builder.InsertCell();
        builder.Writeln("B");
        builder.InsertCell();
        builder.Writeln("C");
        builder.InsertCell();
        builder.Writeln("D");
        builder.EndRow();

        builder.EndTable();

        // Save the document.
        string outputPath = Path.Combine(Environment.CurrentDirectory, "ReportWithMultipleTables.docx");
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new Exception("The output document was not created.");

        Console.WriteLine($"Document successfully created at: {outputPath}");
    }
}
