using System;
using Aspose.Words;
using Aspose.Words.Lists;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Add a numbered list based on the default template.
        List list = doc.Lists.Add(ListTemplate.NumberDefault);

        // Configure the first list level.
        ListLevel level = list.ListLevels[0];
        level.Font.Size = 12;                         // Font size for the number.
        level.NumberStyle = NumberStyle.Arabic;       // Use Arabic numbers.
        level.NumberFormat = "%1.";                   // Number format (e.g., "1.").
        level.TrailingCharacter = ListTrailingCharacter.Tab; // Place a tab after the number.
        level.NumberPosition = -18;                   // Position of the number (negative moves it left).
        level.TextPosition = 36;                      // Position where the text starts.
        level.TabPosition = 36;                       // Tab stop aligns the text after the number.

        // Use DocumentBuilder to add list items.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.ListFormat.List = list;
        builder.Writeln("First item");
        builder.Writeln("Second item");
        builder.Writeln("Third item");
        builder.ListFormat.RemoveNumbers(); // End the list.

        // Save the document to a file.
        doc.Save("CustomTabStopList.docx");
    }
}
