using System;
using Aspose.Words;
using Aspose.Words.Lists;

namespace ListTabPositionExample
{
    public class Program
    {
        public static void Main()
        {
            // Create a new empty document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Add a numbered list using a built‑in template.
            List list = doc.Lists.Add(ListTemplate.NumberDefault);

            // Configure the first level of the list.
            // Set TabPosition to 72 points (1 inch) so that text aligns after the number.
            ListLevel level = list.ListLevels[0];
            level.TabPosition = 72; // points

            // Apply the list to the builder.
            builder.ListFormat.List = list;

            // Write list items.
            builder.Writeln("First item");
            builder.Writeln("Second item");
            builder.Writeln("Third item");

            // Remove list formatting after the list is complete.
            builder.ListFormat.RemoveNumbers();

            // Save the document to a file.
            doc.Save("ListTabPosition.docx");
        }
    }
}
