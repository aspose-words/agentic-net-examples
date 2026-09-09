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

        // Add a numbered list based on the default template.
        List list = doc.Lists.Add(ListTemplate.NumberDefault);

        // Configure the first list level.
        ListLevel level0 = list.ListLevels[0];
        level0.Font.Color = System.Drawing.Color.DarkBlue;
        level0.Font.Size = 12;
        level0.NumberStyle = NumberStyle.Arabic;
        level0.NumberFormat = "%1.";
        level0.StartAt = 1;

        // Position of the number (negative moves it left of the left indent).
        level0.NumberPosition = -18;   // points
        // Position where the text of the paragraph starts.
        level0.TextPosition = 36;      // points
        // Set a tab stop after the number so that the text aligns after the tab.
        level0.TabPosition = 36;       // points
        // Use a tab character as the separator between number and text.
        level0.TrailingCharacter = ListTrailingCharacter.Tab;

        // Apply the list to subsequent paragraphs.
        builder.ListFormat.List = list;

        // First level items.
        builder.Writeln("First level item 1");
        builder.Writeln("First level item 2");

        // Increase list level (second level).
        builder.ListFormat.ListIndent();
        builder.Writeln("Second level item 1");
        builder.Writeln("Second level item 2");

        // Decrease back to first level.
        builder.ListFormat.ListOutdent();
        builder.Writeln("First level item 3");

        // Remove list formatting.
        builder.ListFormat.RemoveNumbers();

        // Save the document to disk.
        doc.Save("CustomTabList.docx");
    }
}
