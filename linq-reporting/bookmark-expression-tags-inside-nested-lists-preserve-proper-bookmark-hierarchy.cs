using System;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Words.Lists;

class Program
{
    static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Create a numbered list that will hold the nested items.
        List list = doc.Lists.Add(ListTemplate.NumberDefault);

        // -------------------------------------------------
        // First level list item – bookmark: Item1
        // -------------------------------------------------
        builder.ListFormat.List = list;          // Apply the list to the builder.
        builder.ListFormat.ListLevelNumber = 0;  // Level 0 = first level.
        builder.StartBookmark("Item1");          // <<bookmark Item1>>
        builder.Writeln("First level item");
        builder.EndBookmark("Item1");            // <</bookmark Item1>>

        // -------------------------------------------------
        // Second level list item – bookmark: Item1_1
        // -------------------------------------------------
        builder.ListFormat.ListLevelNumber = 1;  // Level 1 = second level.
        builder.StartBookmark("Item1_1");        // <<bookmark Item1_1>>
        builder.Writeln("Second level item");
        builder.EndBookmark("Item1_1");          // <</bookmark Item1_1>>

        // -------------------------------------------------
        // Third level list item – bookmark: Item1_1_1
        // -------------------------------------------------
        builder.ListFormat.ListLevelNumber = 2;  // Level 2 = third level.
        builder.StartBookmark("Item1_1_1");      // <<bookmark Item1_1_1>>
        builder.Writeln("Third level item");
        builder.EndBookmark("Item1_1_1");        // <</bookmark Item1_1_1>>

        // Remove list formatting so subsequent paragraphs are not part of the list.
        builder.ListFormat.List = null;

        // -------------------------------------------------
        // Configure outline levels for the bookmarks.
        // When the document is saved as PDF the bookmarks will appear
        // in the navigation pane with the same hierarchical structure.
        // -------------------------------------------------
        PdfSaveOptions pdfOptions = new PdfSaveOptions();
        BookmarksOutlineLevelCollection outline = pdfOptions.OutlineOptions.BookmarksOutlineLevels;
        outline.Add("Item1", 1);        // Level 1 in the PDF outline.
        outline.Add("Item1_1", 2);      // Level 2 in the PDF outline.
        outline.Add("Item1_1_1", 3);    // Level 3 in the PDF outline.

        // Save the document as PDF.
        doc.Save("NestedListWithBookmarks.pdf", pdfOptions);
    }
}
