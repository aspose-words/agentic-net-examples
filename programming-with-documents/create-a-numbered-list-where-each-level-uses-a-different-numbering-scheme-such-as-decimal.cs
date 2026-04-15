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

            // Create a DocumentBuilder which will be used to insert content.
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Add a multilevel list based on the default numbered template.
            List list = doc.Lists.Add(ListTemplate.NumberDefault);

            // Level 0 – decimal numbers (1., 2., 3., ...).
            ListLevel level0 = list.ListLevels[0];
            level0.NumberStyle = NumberStyle.Arabic;
            level0.NumberFormat = "\x0000"; // Placeholder for the number.
            level0.NumberPosition = -18;    // Position of the number.
            level0.TextPosition = 36;       // Position of the text after the number.

            // Level 1 – lower‑roman numbers (i., ii., iii., ...).
            ListLevel level1 = list.ListLevels[1];
            level1.NumberStyle = NumberStyle.LowercaseRoman;
            level1.NumberFormat = "\x0000";
            level1.NumberPosition = -18;
            level1.TextPosition = 36;

            // Apply the list to the builder.
            builder.ListFormat.List = list;

            // First level items (decimal).
            builder.ListFormat.ListLevelNumber = 0;
            builder.Writeln("First level item 1");
            builder.Writeln("First level item 2");

            // Indent to second level (lower‑roman).
            builder.ListFormat.ListIndent();
            builder.Writeln("Second level item i");
            builder.Writeln("Second level item ii");

            // Return to first level.
            builder.ListFormat.ListOutdent();
            builder.Writeln("First level item 3");

            // Remove list formatting from subsequent paragraphs.
            builder.ListFormat.RemoveNumbers();

            // Ensure the output directory exists.
            string outputDir = Path.Combine(Environment.CurrentDirectory, "Output");
            Directory.CreateDirectory(outputDir);

            // Save the document.
            string outputPath = Path.Combine(outputDir, "NumberedListDifferentSchemes.docx");
            doc.Save(outputPath);
        }
    }
}
