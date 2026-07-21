using System;
using Aspose.Words;
using Aspose.Words.Lists;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Create a list based on the built‑in numbered template.
        List list = doc.Lists.Add(ListTemplate.NumberDefault);

        // Access the first list level (level 0) and configure its tab position.
        ListLevel level = list.ListLevels[0];
        // Set the tab position to 72 points (1 inch) so that the text aligns after the number.
        level.TabPosition = 72;
        // Ensure the separator after the number is a tab character.
        level.TrailingCharacter = ListTrailingCharacter.Tab;

        // Use DocumentBuilder to add some list items that will use the custom formatting.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.ListFormat.List = list;
        builder.Writeln("First item");
        builder.Writeln("Second item");
        builder.Writeln("Third item");
        // Remove list formatting after the example.
        builder.ListFormat.RemoveNumbers();

        // Save the document to the current directory.
        string outputPath = System.IO.Path.Combine(Environment.CurrentDirectory, "ListTabPosition.docx");
        doc.Save(outputPath);
    }
}
