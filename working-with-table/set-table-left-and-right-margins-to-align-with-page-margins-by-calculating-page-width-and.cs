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

        // Optionally set custom page margins (in points).
        // 1 inch = 72 points.
        doc.FirstSection.PageSetup.LeftMargin = 72;   // 1 inch
        doc.FirstSection.PageSetup.RightMargin = 72;  // 1 inch

        // Build a simple 2x2 table.
        Table table = builder.StartTable();
        builder.InsertCell();
        builder.Write("Cell 1");
        builder.InsertCell();
        builder.Write("Cell 2");
        builder.EndRow();

        builder.InsertCell();
        builder.Write("Cell 3");
        builder.InsertCell();
        builder.Write("Cell 4");
        builder.EndRow();
        builder.EndTable();

        // Calculate the usable page width (page width minus left and right margins).
        PageSetup setup = doc.FirstSection.PageSetup;
        double usablePageWidth = setup.PageWidth - setup.LeftMargin - setup.RightMargin;

        // Align the table with the page margins:
        // - Set the left indent of the table to the left margin.
        // - Set the preferred width of the table to the usable page width.
        table.LeftIndent = setup.LeftMargin;
        table.PreferredWidth = PreferredWidth.FromPoints(usablePageWidth);

        // Save the document.
        string outputPath = "TableMarginsAligned.docx";
        doc.Save(outputPath);

        // Simple validation to ensure the file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException($"Failed to create the output file: {outputPath}");
    }
}
