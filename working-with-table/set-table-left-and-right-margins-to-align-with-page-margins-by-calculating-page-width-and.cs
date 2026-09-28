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

        // Build a simple 2x2 table.
        builder.StartTable();
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

        // Retrieve the created table.
        Table table = (Table)doc.GetChild(NodeType.Table, 0, true);

        // Calculate page usable width based on page size and margins.
        PageSetup pageSetup = doc.FirstSection.PageSetup;
        double pageWidth = pageSetup.PageWidth;          // Total page width (points).
        double leftMargin = pageSetup.LeftMargin;       // Left margin (points).
        double rightMargin = pageSetup.RightMargin;     // Right margin (points).

        // Align table left edge with the left page margin.
        table.LeftIndent = leftMargin;

        // Set table width so its right edge aligns with the right page margin.
        double usableWidth = pageWidth - leftMargin - rightMargin;
        table.PreferredWidth = PreferredWidth.FromPoints(usableWidth);

        // Save the document.
        string outputPath = "TableMargins.docx";
        doc.Save(outputPath);

        // Verify that the file was saved.
        if (!File.Exists(outputPath))
            throw new Exception("Document was not saved correctly.");
    }
}
