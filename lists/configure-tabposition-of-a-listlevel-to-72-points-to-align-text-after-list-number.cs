using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Lists;

namespace ListTabPositionExample
{
    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();

            // Add a list based on the default numbered template.
            List list = doc.Lists.Add(ListTemplate.NumberDefault);

            // Configure the first list level.
            ListLevel level = list.ListLevels[0];
            // Set the trailing character to a tab so that TabPosition takes effect.
            level.TrailingCharacter = ListTrailingCharacter.Tab;
            // Set the tab position to 72 points (1 inch) to align the text after the number.
            level.TabPosition = 72;
            // Optional: adjust other positions for better appearance.
            level.NumberPosition = -36;   // Position of the number.
            level.TextPosition = 108;    // Position of the text after the tab.

            // Use DocumentBuilder to add list items.
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.ListFormat.List = list;
            builder.Writeln("First item with custom tab position.");
            builder.Writeln("Second item with custom tab position.");

            // Remove list formatting from subsequent paragraphs.
            builder.ListFormat.RemoveNumbers();

            // Save the document to the current directory.
            string outputPath = Path.Combine(Environment.CurrentDirectory, "ListTabPosition.docx");
            doc.Save(outputPath);
        }
    }
}
