using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

public class SplitDocumentWithCompleteTableRows
{
    public static void Main()
    {
        // Create a sample document with a table that spans multiple pages.
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);

        // Set page size to A4 to ensure the table will need several pages.
        sourceDoc.FirstSection.PageSetup.PaperSize = PaperSize.A4;

        // Insert a table with many rows.
        Table table = builder.StartTable();

        // Add a header row.
        builder.InsertCell();
        builder.Write("Header");
        builder.EndRow();

        // Add rows that will force page breaks.
        for (int i = 1; i <= 30; i++)
        {
            builder.InsertCell();
            // Write enough text to make the row relatively tall.
            builder.Write($"Row {i} - Lorem ipsum dolor sit amet, consectetur adipiscing elit. " +
                          "Sed do eiusmod tempor incididunt ut labore et dolore magna aliqua.");
            builder.EndRow();

            // Prevent the row from being split across pages.
            Row currentRow = table.LastRow;
            currentRow.RowFormat.AllowBreakAcrossPages = false;
        }

        builder.EndTable();

        // Save the source document.
        string sourcePath = "Source.docx";
        sourceDoc.Save(sourcePath);

        // Load the document again (simulating a real scenario).
        Document doc = new Document(sourcePath);

        // Ensure the document has more than one page.
        int pageCount = doc.PageCount;
        if (pageCount < 2)
            throw new InvalidOperationException("The document does not span multiple pages as expected.");

        // Split the document by pages, preserving whole table rows.
        for (int page = 1; page <= pageCount; page++)
        {
            // Extract a single page. Use zero‑based start index and a count of 1.
            Document part = doc.ExtractPages(page - 1, 1);
            string partPath = $"Part_{page}.docx";
            part.Save(partPath);

            // Validate that the split part was created.
            if (!File.Exists(partPath))
                throw new FileNotFoundException($"Failed to create split part: {partPath}");
        }

        // All split parts have been created successfully.
        Console.WriteLine($"Document split into {pageCount} parts with complete table rows.");
    }
}
