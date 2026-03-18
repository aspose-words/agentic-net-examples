using System;
using Aspose.Words;
using Aspose.Words.Lists;
using Aspose.Words.Drawing;
using System.Drawing;

class ApplyCustomTabStopToList
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Create a list based on the built‑in numbered template.
        List list = doc.Lists.Add(ListTemplate.NumberDefault);

        // Access the first list level (level 0).
        ListLevel level = list.ListLevels[0];

        // Use a tab character as the separator between the number and the paragraph text.
        level.TrailingCharacter = ListTrailingCharacter.Tab;

        // Set the position of the tab stop (in points). 144 points = 2 inches.
        level.TabPosition = 144.0;

        // Optionally adjust the number and text positions so the number aligns correctly.
        // NumberPosition is the left edge of the number/bullet.
        level.NumberPosition = -36.0;   // Move the number left of the left indent.
        // TextPosition is the left edge of the paragraph text after the tab.
        level.TextPosition = 144.0;     // Align text with the tab stop.

        // Use a DocumentBuilder to add list items.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.ListFormat.List = list; // Start using the custom list.

        builder.Writeln("First item – the text starts after the tab.");
        builder.Writeln("Second item – also aligned after the tab.");

        // Add a sub‑level to demonstrate that the custom tab stop only applies to level 0.
        builder.ListFormat.ListIndent(); // Increase list level to 1.
        builder.Writeln("Sub‑item – uses default tab settings.");

        // Return to the top level.
        builder.ListFormat.ListOutdent();

        // Remove list formatting from subsequent paragraphs.
        builder.ListFormat.RemoveNumbers();

        // Save the document.
        doc.Save("CustomListTabPosition.docx");
    }
}
