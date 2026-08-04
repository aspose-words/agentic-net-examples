using System;
using Aspose.Words;
using Aspose.Words.Tables;

public class OptimizeTableRendering
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Start a table.
        Table table = builder.StartTable();

        // Insert the first cell of the first row (required before any formatting).
        builder.InsertCell();

        // Turn off automatic layout updates by postponing the layout refresh.
        // Aspose.Words updates the layout lazily; we will force a single layout update
        // after all rows have been added to avoid repeated recalculations.
        // (No explicit property to disable layout; we simply avoid calling UpdatePageLayout
        // until the batch operation is finished.)

        // Add a large number of rows to the table.
        const int rowCount = 5000; // Example large number for performance testing.
        for (int i = 0; i < rowCount; i++)
        {
            // First cell of the row.
            builder.InsertCell();
            builder.Write($"Row {i + 1}, Cell 1");

            // Second cell of the row.
            builder.InsertCell();
            builder.Write($"Row {i + 1}, Cell 2");

            // End the current row.
            builder.EndRow();
        }

        // End the table construction.
        builder.EndTable();

        // After all modifications are done, update the layout once.
        doc.UpdatePageLayout();

        // Save the document to verify the result.
        const string outputPath = "OptimizedTable.docx";
        doc.Save(outputPath);

        Console.WriteLine($"Document saved to {outputPath}");
    }
}
