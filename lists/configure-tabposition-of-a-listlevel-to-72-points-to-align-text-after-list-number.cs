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

            // Add a numbered list based on the default template.
            List list = doc.Lists.Add(ListTemplate.NumberDefault);

            // Configure the first list level.
            ListLevel level = list.ListLevels[0];
            // Use a tab as the trailing character so that TabPosition takes effect.
            level.TrailingCharacter = ListTrailingCharacter.Tab;
            // Set the tab position to 72 points (1 inch) to align the text after the number.
            level.TabPosition = 72.0;
            // Optional: adjust other positions for clarity.
            level.NumberPosition = -36.0;
            level.TextPosition = 144.0;

            // Build list items using DocumentBuilder.
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.ListFormat.List = list;
            builder.Writeln("First item");
            builder.Writeln("Second item");
            builder.Writeln("Third item");
            builder.ListFormat.RemoveNumbers();

            // Save the document to the current directory.
            string outputPath = Path.Combine(Environment.CurrentDirectory, "ListTabPosition.docx");
            doc.Save(outputPath);
        }
    }
}
