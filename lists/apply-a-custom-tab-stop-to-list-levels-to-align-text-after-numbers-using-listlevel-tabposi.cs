using System;
using Aspose.Words;
using Aspose.Words.Lists;

public class Program
{
    public static void Main()
    {
        // Create a new document and a DocumentBuilder.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Create a numbered list using a built‑in template.
        List list = doc.Lists.Add(ListTemplate.NumberDefault);

        // Configure the first level of the list.
        ListLevel level = list.ListLevels[0];
        // Set a custom tab stop (1 inch = 72 points) to align text after the number.
        level.TabPosition = 72;
        // Optional adjustments for better appearance.
        level.NumberPosition = 36;
        level.Alignment = ListLevelAlignment.Right;

        // Apply the list to the builder.
        builder.ListFormat.List = list;
        builder.ListFormat.ListLevelNumber = 0; // First level.

        // Write list items.
        builder.Writeln("First item with custom tab stop.");
        builder.Writeln("Second item with custom tab stop.");
        builder.Writeln("Third item with custom tab stop.");

        // End the list.
        builder.ListFormat.RemoveNumbers();

        // Save the document.
        doc.Save("ListWithCustomTab.docx");
    }
}
