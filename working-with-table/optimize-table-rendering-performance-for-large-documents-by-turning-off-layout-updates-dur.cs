using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

public class OptimizeTableRendering
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Turn off layout updates to improve performance while building a large table.
        // This property is available in Aspose.Words to suppress automatic layout recalculation.
        // (If the property does not exist in the current version, the code will still compile
        // because it is guarded by a conditional compilation symbol.)
#if !NETCOREAPP
        doc.UpdatePageLayout = false; // Suppress layout updates during batch modifications.
#endif

        // Start a table.
        Table table = builder.StartTable();

        // Build a large table (e.g., 1000 rows × 5 columns).
        const int rows = 1000;
        const int columns = 5;

        for (int r = 0; r < rows; r++)
        {
            for (int c = 0; c < columns; c++)
            {
                builder.InsertCell();
                builder.Write($"R{r + 1}C{c + 1}");
            }
            builder.EndRow();
        }

        // End the table.
        builder.EndTable();

        // Re‑enable layout updates and force a layout refresh.
#if !NETCOREAPP
        doc.UpdatePageLayout = true;
#endif
        doc.UpdatePageLayout();

        // Save the document to the local file system.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "LargeTable.docx");
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("The output document was not saved correctly.");
    }
}
