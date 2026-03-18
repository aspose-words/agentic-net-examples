using System;
using Aspose.Words;
using Aspose.Words.Lists;

class Program
{
    static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add a multilevel list based on the built‑in NumberDefault template.
        List list = doc.Lists.Add(ListTemplate.NumberDefault);

        // -----------------------------------------------------------------
        // Level 0 – decimal (Arabic) numbering, e.g. 1., 2., 3.
        // -----------------------------------------------------------------
        ListLevel level0 = list.ListLevels[0];
        level0.NumberStyle = NumberStyle.Arabic;               // Arabic numbers.
        level0.NumberFormat = "\x0000.";                       // "\x0000" is the placeholder for the current level number.

        // -----------------------------------------------------------------
        // Level 1 – lower‑roman numbering, e.g. i., ii., iii.
        // -----------------------------------------------------------------
        ListLevel level1 = list.ListLevels[1];
        level1.NumberStyle = NumberStyle.LowercaseRoman;       // Lower‑case Roman numerals.
        level1.NumberFormat = "\x0000.";                       // Same placeholder, but the style changes the appearance.

        // Apply the list to the first‑level items.
        builder.ListFormat.List = list;
        builder.Writeln("First level item 1");
        builder.Writeln("First level item 2");

        // Increase the list level to produce second‑level items.
        builder.ListFormat.ListIndent();
        builder.Writeln("Second level item i");
        builder.Writeln("Second level item ii");

        // Return to the first level.
        builder.ListFormat.ListOutdent();
        builder.Writeln("First level item 3");

        // End list formatting for any subsequent paragraphs.
        builder.ListFormat.RemoveNumbers();

        // Save the document to disk.
        doc.Save("NumberedListDifferentStyles.docx");
    }
}
