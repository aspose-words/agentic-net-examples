using System;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Drawing.Charts; // Needed for chart detection

class AddBookmarksToListItems
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // -----------------------------------------------------------------
        // 1. Build a numbered list with several items.
        // -----------------------------------------------------------------
        builder.ListFormat.ApplyNumberDefault(); // start numbering
        for (int i = 1; i <= 5; i++)
        {
            // Insert a paragraph that is part of the list.
            builder.Writeln($"List item {i}");
        }
        builder.ListFormat.RemoveNumbers(); // stop list formatting

        // -----------------------------------------------------------------
        // 2. Iterate through all paragraphs that belong to the list.
        //    For each paragraph that does NOT contain a chart shape,
        //    insert a bookmark that surrounds the whole paragraph.
        // -----------------------------------------------------------------
        int bookmarkIndex = 1;
        foreach (Paragraph para in doc.FirstSection.Body.Paragraphs)
        {
            // Ensure the paragraph is a list item.
            if (!para.IsListItem) continue;

            // Detect whether the paragraph (or any of its descendants) contains a chart.
            bool containsChart = false;
            NodeCollection shapes = para.GetChildNodes(NodeType.Shape, true);
            foreach (Shape shape in shapes)
            {
                // In Aspose.Words a chart is stored as a Shape with HasChart == true.
                if (shape.HasChart)
                {
                    containsChart = true;
                    break;
                }
            }

            // Skip bookmark creation if a chart is present.
            if (containsChart) continue;

            // Move the builder's cursor to the start of the paragraph.
            builder.MoveTo(para);
            string bookmarkName = $"ListItem_{bookmarkIndex}_BM";

            // Insert the bookmark start and end around the paragraph.
            builder.StartBookmark(bookmarkName);
            builder.EndBookmark(bookmarkName);

            bookmarkIndex++;
        }

        // -----------------------------------------------------------------
        // 3. Save the document.
        // -----------------------------------------------------------------
        doc.Save("ListItemsWithBookmarks.docx");
    }
}
