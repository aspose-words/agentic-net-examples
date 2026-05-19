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

        // Add a list based on the default numbered template.
        List list = doc.Lists.Add(ListTemplate.NumberDefault);

        // Configure the first level of the list.
        ListLevel level = list.ListLevels[0];
        // Use a simple numeric format.
        level.NumberFormat = "%1.";
        // Place a tab character after the number.
        level.TrailingCharacter = ListTrailingCharacter.Tab;
        // Position of the number (negative moves it left of the left indent).
        level.NumberPosition = -36; // points
        // Position where the text starts after the tab.
        level.TextPosition = 144; // points
        // Custom tab stop that aligns the text after the number.
        level.TabPosition = 144; // points

        // Use DocumentBuilder to add list items.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.ListFormat.List = list;
        builder.Writeln("First item with custom tab stop.");
        builder.Writeln("Second item with custom tab stop.");

        // Save the document to the current directory.
        string outputPath = Path.Combine(Environment.CurrentDirectory, "CustomTabStopList.docx");
        doc.Save(outputPath);
    }
}
