using System;
using Aspose.Words;
using Aspose.Words.Lists;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Create a new list based on a built‑in template (NumberDefault) and configure its first level.
        List list = doc.Lists.Add(ListTemplate.NumberDefault);
        ListLevel level = list.ListLevels[0];
        level.NumberStyle = NumberStyle.Arabic;          // Use Arabic numerals.
        level.NumberFormat = "%1.";                      // Format like "1."
        level.Alignment = ListLevelAlignment.Left;       // Align numbers to the left.
        level.TabPosition = 72;                          // Set tab position to 72 points (1 inch).
        level.NumberPosition = 0;                        // Position of the number.
        level.TextPosition = 72;                         // Position of the text after the tab.

        // Apply the list to several paragraphs.
        builder.ListFormat.List = list;
        builder.Writeln("First item");
        builder.Writeln("Second item");
        builder.Writeln("Third item");
        builder.ListFormat.RemoveNumbers(); // Stop list formatting.

        // Save the document to a file.
        doc.Save("ListWithTabPosition.docx");
    }
}
