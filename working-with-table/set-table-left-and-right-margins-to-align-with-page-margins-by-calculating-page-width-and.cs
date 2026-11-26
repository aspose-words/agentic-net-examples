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

        // Build a simple 3x3 table.
        builder.StartTable();
        for (int row = 0; row < 3; row++)
        {
            for (int col = 0; col < 3; col++)
            {
                builder.InsertCell();
                builder.Writeln($"R{row + 1}C{col + 1}");
            }
            builder.EndRow();
        }
        builder.EndTable();

        // Retrieve the created table (first table in the document).
        Table table = (Table)doc.GetChildNodes(NodeType.Table, true)[0];

        // Get page setup values (in points).
        PageSetup pageSetup = doc.FirstSection.PageSetup;
        double pageWidth = pageSetup.PageWidth;
        double leftMargin = pageSetup.LeftMargin;
        double rightMargin = pageSetup.RightMargin;

        // Align table left edge with left margin.
        table.LeftIndent = leftMargin;

        // Set table width so that its right edge aligns with the right margin.
        double tableWidth = pageWidth - leftMargin - rightMargin;
        table.PreferredWidth = PreferredWidth.FromPoints(tableWidth);

        // Save the document.
        string outputPath = "TableMargins.docx";
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("The output document was not created.");

        // End of example.
    }
}
