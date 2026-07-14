using System;
using Aspose.Words;
using Aspose.Words.Lists;

public class Program
{
    public static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();

        // Add a numbered list based on the default template.
        List list = doc.Lists.Add(ListTemplate.NumberDefault);

        // ----- Configure the first list level -----
        ListLevel level0 = list.ListLevels[0];
        // Use the default number format (e.g., "1.") – the placeholder "\x0000" represents the number.
        level0.NumberFormat = "\x0000.";
        // Position of the number relative to the left margin (negative moves it left of the text indent).
        level0.NumberPosition = -18; // points
        // Position where the paragraph text starts after the number.
        level0.TextPosition = 36; // points
        // Custom tab stop that aligns the text after the number.
        level0.TabPosition = 36; // points (same as TextPosition)
        // Insert a tab character between the number and the text.
        level0.TrailingCharacter = ListTrailingCharacter.Tab;

        // ----- Configure the second list level (optional) -----
        ListLevel level1 = list.ListLevels[1];
        level1.NumberFormat = "\x0000.\x0001.";
        level1.NumberPosition = -18;
        level1.TextPosition = 72;
        level1.TabPosition = 72;
        level1.TrailingCharacter = ListTrailingCharacter.Tab;

        // Use DocumentBuilder to add list items.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.ListFormat.List = list;

        // First‑level items.
        builder.Writeln("First level item 1");
        builder.Writeln("First level item 2");

        // Indent to second level and add items.
        builder.ListFormat.ListIndent();
        builder.Writeln("Second level item 1");
        builder.Writeln("Second level item 2");
        builder.ListFormat.ListOutdent();

        // Remove list formatting from subsequent paragraphs.
        builder.ListFormat.RemoveNumbers();

        // Save the document to the current directory.
        doc.Save("CustomList.docx");
    }
}
