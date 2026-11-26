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

        // Build a simple 1x1 table.
        builder.StartTable();
        builder.InsertCell();
        builder.Writeln("Sample cell");
        builder.EndRow();
        builder.EndTable();

        // Retrieve the created table.
        Table table = (Table)doc.GetChild(NodeType.Table, 0, true);

        // Convert 1 centimeter to points (1 cm = 28.3465 points).
        const double cmToPoints = 28.3465;

        // Set left indent of the table.
        table.LeftIndent = cmToPoints;

        // Aspose.Words Table does not expose a RightIndent property.
        // To achieve a right margin of 1 cm, set the table width so that
        // the remaining space on the page equals the desired right indent.
        double pageContentWidth = doc.FirstSection.PageSetup.PageWidth
                                   - doc.FirstSection.PageSetup.LeftMargin
                                   - doc.FirstSection.PageSetup.RightMargin;

        double tableWidth = pageContentWidth - (2 * cmToPoints); // left + right margins
        table.PreferredWidth = PreferredWidth.FromPoints(tableWidth);

        // Save the document to disk.
        const string outputFile = "TableIndent.docx";
        doc.Save(outputFile);

        // Verify that the file was created.
        if (!File.Exists(outputFile))
            throw new Exception("Failed to create the output document.");
    }
}
