using System;
using Aspose.Words;
using Aspose.Words.Lists;

namespace MultiLevelListExample
{
    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Create a custom list based on a generic template (NumberDefault) that we will fully customize.
            List customList = doc.Lists.Add(ListTemplate.NumberDefault);

            // Level 0 – Uppercase Roman numerals (e.g., I., II., III.)
            ListLevel level0 = customList.ListLevels[0];
            level0.NumberStyle = NumberStyle.UppercaseRoman;
            level0.NumberFormat = "%1.";
            level0.Font.Name = "Times New Roman";
            level0.Font.Size = 12;

            // Level 1 – Bullet.
            ListLevel level1 = customList.ListLevels[1];
            level1.NumberStyle = NumberStyle.Bullet;
            level1.NumberFormat = "•";
            level1.Font.Name = "Arial";
            level1.Font.Size = 12;

            // Level 2 – Lowercase alphabetic (e.g., a), b), c))
            ListLevel level2 = customList.ListLevels[2];
            level2.NumberStyle = NumberStyle.LowercaseLetter;
            level2.NumberFormat = "%3)";
            level2.Font.Name = "Courier New";
            level2.Font.Size = 12;

            // Apply the custom list to the builder.
            builder.ListFormat.List = customList;

            // Insert a level‑0 item.
            builder.ListFormat.ListLevelNumber = 0;
            builder.Writeln("First level item");

            // Insert a level‑1 item.
            builder.ListFormat.ListLevelNumber = 1;
            builder.Writeln("Second level bullet item");

            // Insert a level‑2 item.
            builder.ListFormat.ListLevelNumber = 2;
            builder.Writeln("Third level alphabetic item");

            // End list formatting.
            builder.ListFormat.RemoveNumbers();

            // Save the document.
            string outputPath = "MultiLevelList.docx";
            doc.Save(outputPath);
        }
    }
}
