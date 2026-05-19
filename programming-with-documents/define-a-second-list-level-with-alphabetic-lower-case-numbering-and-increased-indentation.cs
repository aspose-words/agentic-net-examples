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

        // Add a multilevel numbered list based on the default template.
        // The default template already defines lower‑case letters for the second level.
        List list = doc.Lists.Add(ListTemplate.NumberDefault);
        builder.ListFormat.List = list;

        // Customize the second list level (index 1) to increase its indentation.
        ListLevel secondLevel = list.ListLevels[1];
        secondLevel.NumberStyle = NumberStyle.LowercaseLetter; // Ensure lower‑case letters.
        secondLevel.NumberPosition = -36;   // Move the number leftwards.
        secondLevel.TextPosition = 144;    // Increase the text indent.
        secondLevel.TabPosition = 144;     // Align tab for the second level.

        // First‑level items.
        builder.Writeln("First level item 1");
        builder.Writeln("First level item 2");

        // Increase the list level to the second level.
        builder.ListFormat.ListIndent();

        // Second‑level items will be numbered a., b., c. with the increased indent.
        builder.Writeln("Second level item a");
        builder.Writeln("Second level item b");
        builder.Writeln("Second level item c");

        // Return to the first level.
        builder.ListFormat.ListOutdent();

        // Additional first‑level items.
        builder.Writeln("First level item 3");

        // End the list.
        builder.ListFormat.RemoveNumbers();

        // Save the document to the current directory.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "SecondLevelList.docx");
        doc.Save(outputPath);
    }
}
