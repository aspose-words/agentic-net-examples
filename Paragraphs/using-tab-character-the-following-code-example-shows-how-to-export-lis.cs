using System;
using Aspose.Words;
using Aspose.Words.Lists;

class ExportListWithTabs
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add a numbered list based on the built‑in template.
        List list = doc.Lists.Add(ListTemplate.NumberDefault);

        // Configure every level of the list to place a TAB after the list label.
        foreach (ListLevel level in list.ListLevels)
        {
            level.TrailingCharacter = ListTrailingCharacter.Tab;
        }

        // Apply the list to the builder and add some items.
        builder.ListFormat.List = list;
        builder.Writeln("Item 1");               // Level 0
        builder.ListFormat.ListIndent();         // Move to Level 1
        builder.Writeln("Item 2");               // Level 1
        builder.ListFormat.ListIndent();         // Move to Level 2
        builder.Writeln("Item 3");               // Level 2
        builder.ListFormat.ListOutdent();        // Back to Level 1
        builder.Writeln("Back to level 1");      // Level 1
        builder.ListFormat.RemoveNumbers();      // End the list

        // Save the document as DOCX.
        doc.Save("ExportListWithTabs.docx");
    }
}
