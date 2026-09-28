using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

public class OptimizeTableRendering
{
    public static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();

        // Use DocumentBuilder to construct the table.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Define table dimensions.
        const int rowCount = 500;   // Adjust as needed for a large document.
        const int columnCount = 10;

        // Start the table.
        builder.StartTable();

        // Populate the table with sample data.
        for (int row = 0; row < rowCount; row++)
        {
            for (int col = 0; col < columnCount; col++)
            {
                builder.InsertCell();
                builder.Writeln($"R{row + 1}C{col + 1}");
            }

            // End the current row.
            builder.EndRow();
        }

        // End the table.
        builder.EndTable();

        // Force a layout refresh after all modifications are done.
        doc.UpdatePageLayout();

        // Save the document to disk.
        string outputPath = "LargeTable.docx";
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException($"Failed to create the output file: {outputPath}");
    }
}
