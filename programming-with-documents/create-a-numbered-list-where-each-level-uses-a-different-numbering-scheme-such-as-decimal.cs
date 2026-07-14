using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Lists;

namespace AsposeWordsListExample
{
    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Create a multilevel list based on the default numbered template.
            List list = doc.Lists.Add(ListTemplate.NumberDefault);

            // Level 0 – decimal numbers (1., 2., 3., ...).
            ListLevel level0 = list.ListLevels[0];
            level0.NumberStyle = NumberStyle.Arabic;
            level0.NumberFormat = "%1.";
            level0.StartAt = 1;

            // Level 1 – lower‑roman numbers (i., ii., iii., ...).
            ListLevel level1 = list.ListLevels[1];
            level1.NumberStyle = NumberStyle.LowercaseRoman;
            level1.NumberFormat = "%2.";
            level1.StartAt = 1;

            // Apply the list to the builder.
            builder.ListFormat.List = list;

            // First‑level items.
            builder.ListFormat.ListLevelNumber = 0;
            builder.Writeln("First level item 1");
            builder.Writeln("First level item 2");

            // Indent to second level.
            builder.ListFormat.ListIndent();

            // Second‑level items.
            builder.Writeln("Second level item 1");
            builder.Writeln("Second level item 2");

            // Outdent back to first level.
            builder.ListFormat.ListOutdent();

            // More first‑level items.
            builder.Writeln("First level item 3");
            builder.Writeln("First level item 4");

            // End the list.
            builder.ListFormat.RemoveNumbers();

            // Determine an output path in the current directory.
            string outputPath = Path.Combine(Environment.CurrentDirectory, "NumberedList.docx");
            doc.Save(outputPath);
        }
    }
}
