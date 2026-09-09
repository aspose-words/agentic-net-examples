using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Lists;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Start a default numbered list (1., 2., 3., ...).
        builder.ListFormat.ApplyNumberDefault();

        // Add a first‑level item.
        builder.Writeln("First level item 1");

        // Increase the list level to create a second level.
        builder.ListFormat.ListIndent();

        // Retrieve the list that is currently applied.
        List list = builder.ListFormat.List;

        // Configure the second level (index 1) to use lower‑case alphabetic numbering
        // and increase its indentation.
        ListLevel secondLevel = list.ListLevels[1];
        secondLevel.NumberStyle = NumberStyle.LowercaseLetter; // a., b., c., ...
        // Increase indentation: move the number left and the text right.
        secondLevel.NumberPosition = -36;   // Position of the number (points).
        secondLevel.TextPosition = 144;    // Position where the text starts (points).
        secondLevel.TabPosition = 144;     // Tab stop for the level (points).

        // Add some second‑level items.
        builder.Writeln("Second level item a");
        builder.Writeln("Second level item b");
        builder.Writeln("Second level item c");

        // Return to the first level.
        builder.ListFormat.ListOutdent();

        // Add another first‑level item.
        builder.Writeln("First level item 2");

        // End the list formatting.
        builder.ListFormat.RemoveNumbers();

        // Save the document to the current directory.
        string outputPath = Path.Combine(Environment.CurrentDirectory, "SecondLevelList.docx");
        doc.Save(outputPath);
    }
}
