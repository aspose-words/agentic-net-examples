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

        // Add a numbered list to the document.
        List list = doc.Lists.Add(ListTemplate.NumberDefault);

        // Configure the second list level (index 1) to use lower‑case alphabetic numbering.
        ListLevel level2 = list.ListLevels[1];
        level2.NumberStyle = NumberStyle.LowercaseLetter;   // a., b., c., …
        level2.NumberFormat = "%1.";                       // format of the number
        // Increase indentation for the second level.
        level2.NumberPosition = 72;    // position of the number (points)
        level2.TextPosition = 108;     // position where the text starts (points)
        level2.TabPosition = 108;      // tab stop after the number

        // Apply the list to the builder.
        builder.ListFormat.List = list;

        // First‑level list item.
        builder.ListFormat.ListLevelNumber = 0; // level 0 = first level
        builder.Writeln("First level item");

        // Second‑level list items.
        builder.ListFormat.ListLevelNumber = 1; // level 1 = second level
        builder.Writeln("Second level item a");
        builder.Writeln("Second level item b");

        // End the list.
        builder.ListFormat.RemoveNumbers();

        // Save the document to a file.
        doc.Save("ListExample.docx");
    }
}
