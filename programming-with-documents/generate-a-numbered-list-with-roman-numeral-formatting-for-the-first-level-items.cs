using System;
using Aspose.Words;
using Aspose.Words.Lists;

namespace AsposeWordsRomanList
{
    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();

            // Create a DocumentBuilder which will be used to insert content.
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Add a new list based on the built‑in numbered template.
            List list = doc.Lists.Add(ListTemplate.NumberDefault);

            // Configure the first level (level 0) to use upper‑case Roman numerals.
            ListLevel level0 = list.ListLevels[0];
            level0.NumberStyle = NumberStyle.UppercaseRoman;   // I, II, III, ...
            level0.NumberFormat = "%1.";                       // Append a period after the number.

            // Apply the list to the builder.
            builder.ListFormat.List = list;
            builder.ListFormat.ListLevelNumber = 0; // Ensure we are on the first level.

            // Add some list items.
            builder.Writeln("First item");
            builder.Writeln("Second item");
            builder.Writeln("Third item");
            builder.Writeln("Fourth item");

            // End the list formatting.
            builder.ListFormat.RemoveNumbers();

            // Save the document to the local file system.
            doc.Save("RomanList.docx");
        }
    }
}
