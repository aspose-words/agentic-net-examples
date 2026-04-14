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

        // Access the first level (level 0) of the list.
        ListLevel level = list.ListLevels[0];

        // Set the trailing character to a tab so that TabPosition takes effect.
        level.TrailingCharacter = ListTrailingCharacter.Tab;

        // Configure the tab position to 72 points (1 inch) to align the text after the number.
        level.TabPosition = 72.0;

        // Optionally adjust other positions for better visual layout.
        level.NumberPosition = -36.0;   // Position of the number.
        level.TextPosition = 144.0;     // Position of the text after the tab.

        // Use DocumentBuilder to add list items.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.ListFormat.List = list;
        builder.Writeln("First item with custom tab position.");
        builder.Writeln("Second item with custom tab position.");

        // Remove list formatting from subsequent paragraphs.
        builder.ListFormat.RemoveNumbers();

        // Define an output folder and file name.
        string outputDir = Path.Combine(Environment.CurrentDirectory, "Output");
        Directory.CreateDirectory(outputDir);
        string outputPath = Path.Combine(outputDir, "ListWithTabPosition.docx");

        // Save the document.
        doc.Save(outputPath);
    }
}
