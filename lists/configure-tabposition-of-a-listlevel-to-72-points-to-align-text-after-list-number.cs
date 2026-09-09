using System;
using Aspose.Words;
using Aspose.Words.Lists;

public class Program
{
    public static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();

        // Add a list based on the default numbered template.
        List list = doc.Lists.Add(ListTemplate.NumberDefault);

        // Configure the first list level.
        ListLevel level = list.ListLevels[0];
        // Use a tab as the trailing character so that TabPosition takes effect.
        level.TrailingCharacter = ListTrailingCharacter.Tab;
        // Set the tab position to 72 points (1 inch) to align the text after the number.
        level.TabPosition = 72;
        // Optional: set number and text positions for clearer layout.
        level.NumberPosition = 0;
        level.TextPosition = 72;

        // Add some paragraphs that use the configured list.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.ListFormat.List = list;
        builder.Writeln("First list item");
        builder.Writeln("Second list item");
        builder.ListFormat.RemoveNumbers();

        // Save the document to the current directory.
        doc.Save("Lists.TabPosition.docx");
    }
}
