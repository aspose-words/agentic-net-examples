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

        // Retrieve the created table.
        Table table = (Table)doc.GetChild(NodeType.Table, 0, true);

        // Get page setup information from the first section.
        PageSetup pageSetup = doc.FirstSection.PageSetup;
        // Cast to float because PageWidth, LeftMargin and RightMargin are double in this version.
        float pageWidth = (float)pageSetup.PageWidth;          // Total page width in points.
        float leftMargin = (float)pageSetup.LeftMargin;        // Left margin in points.
        float rightMargin = (float)pageSetup.RightMargin;      // Right margin in points.

        // Calculate the width that fits exactly between the margins.
        float availableWidth = pageWidth - leftMargin - rightMargin;

        // Align the table with page margins.
        table.LeftIndent = leftMargin;                                   // Align left side.
        table.PreferredWidth = PreferredWidth.FromPoints(availableWidth); // Set width to fill the space.

        // Save the document.
        string outputPath = "TableMargins.docx";
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new Exception("Failed to create the output document.");

        // Inform that the process completed.
        Console.WriteLine($"Document saved to '{Path.GetFullPath(outputPath)}'.");
    }
}
