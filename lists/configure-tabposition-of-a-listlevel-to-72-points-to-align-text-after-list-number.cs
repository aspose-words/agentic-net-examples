using System;
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

            // Create a list based on the built‑in numbered template.
            List list = doc.Lists.Add(ListTemplate.NumberDefault);

            // Access the first level of the list.
            ListLevel level = list.ListLevels[0];

            // Set the tab position to 72 points (1 inch) so that the text aligns after the number.
            level.TabPosition = 72;

            // Use a tab character as the trailing character to make the TabPosition effective.
            level.TrailingCharacter = ListTrailingCharacter.Tab;

            // Build some list items to demonstrate the formatting.
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.ListFormat.List = list;
            builder.Writeln("First item");
            builder.Writeln("Second item");
            builder.Writeln("Third item");

            // Remove list formatting from subsequent paragraphs (optional).
            builder.ListFormat.RemoveNumbers();

            // Save the document to the file system.
            doc.Save("ListTabPosition.docx");
        }
    }
}
