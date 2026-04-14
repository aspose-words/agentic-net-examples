using System;
using Aspose.Words;
using Aspose.Words.Lists;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Create a DocumentBuilder to insert content.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add a multilevel list based on the default numbered template.
        List list = doc.Lists.Add(ListTemplate.NumberDefault);

        // Adjust the indentation of the first list level.
        // The ListLevel class does not have an Indentation property.
        // Use NumberPosition, TextPosition, and TabPosition (all in points) to control indentation.
        ListLevel level0 = list.ListLevels[0];
        level0.NumberPosition = 36; // Position of the number/bullet.
        level0.TextPosition = 36;   // Position of the text (left indent).
        level0.TabPosition = 36;    // Tab stop for the list level.

        // Apply the list to the builder and add some items.
        builder.ListFormat.List = list;
        builder.Writeln("Item 1");
        builder.ListFormat.ListIndent(); // Increase list level.
        builder.Writeln("Subitem 1");
        builder.ListFormat.ListOutdent(); // Return to previous level.
        builder.Writeln("Item 2");

        // Remove list formatting.
        builder.ListFormat.RemoveNumbers();

        // Save the document to the output file.
        string outputPath = "ListIndentation.docx";
        doc.Save(outputPath);
    }
}
