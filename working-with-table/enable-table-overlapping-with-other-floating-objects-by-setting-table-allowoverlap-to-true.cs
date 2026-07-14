using System;
using Aspose.Words;
using Aspose.Words.Tables;

namespace TableAllowOverlapExample
{
    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert some preceding text.
            builder.Writeln("This paragraph appears before the floating table.");

            // Start a new table. This will be a floating table because we will set TextWrapping.
            Table table = builder.StartTable();

            // Insert a single cell with some content.
            builder.InsertCell();
            builder.Write("Floating table cell.");

            // End the row and the table.
            builder.EndRow();
            builder.EndTable();

            // Configure the table to wrap text around it.
            table.TextWrapping = TextWrapping.Around;

            // Position the table somewhere on the page.
            table.AbsoluteHorizontalDistance = 100; // points from the paragraph.
            table.AbsoluteVerticalDistance = 20;   // points from the paragraph.

            // Table.AllowOverlap is read‑only and defaults to true, so no assignment is needed.
            // This ensures the floating table can overlap other floating objects.

            // Insert some following text that will flow around the floating table.
            builder.Writeln("This paragraph appears after the floating table and should wrap around it if overlapping is allowed.");

            // Save the document to a file.
            string outputPath = "TableAllowOverlap.docx";
            doc.Save(outputPath);
        }
    }
}
