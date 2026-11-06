using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a sample document with a table whose rows are forced to stay together.
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);

        // Start a table.
        Table table = builder.StartTable();

        // Add multiple rows with enough text to span several pages.
        for (int i = 1; i <= 30; i++)
        {
            // Insert a cell and write long text.
            builder.InsertCell();
            builder.Write($"Row {i} - " + new string('x', 500));

            // End the current row.
            builder.EndRow();

            // Prevent the row from breaking across pages.
            Row currentRow = table.Rows[table.Rows.Count - 1];
            currentRow.RowFormat.AllowBreakAcrossPages = false;
        }

        // End the table.
        builder.EndTable();

        // Ensure layout is up‑to‑date so that PageCount is correct.
        sourceDoc.UpdatePageLayout();

        // Save the source document (optional, for inspection).
        string sourcePath = "Source.docx";
        sourceDoc.Save(sourcePath);

        // Determine the number of pages in the source document.
        int pageCount = sourceDoc.PageCount;

        // Split the document by pages, ensuring rows remain intact.
        // Document.ExtractPages uses a zero‑based page index.
        for (int pageIndex = 0; pageIndex < pageCount; pageIndex++)
        {
            // Extract a single page as a new document.
            Document part = sourceDoc.ExtractPages(pageIndex, 1);

            // Save the split part. Use 1‑based naming for readability.
            string partPath = $"Part_{pageIndex + 1}.docx";
            part.Save(partPath);
        }

        // Validate that each split part was created.
        for (int pageIndex = 0; pageIndex < pageCount; pageIndex++)
        {
            string partPath = $"Part_{pageIndex + 1}.docx";
            if (!File.Exists(partPath))
                throw new FileNotFoundException($"Expected split file not found: {partPath}");
        }

        // All split parts have been created successfully.
        Console.WriteLine($"Document split into {pageCount} parts successfully.");
    }
}
