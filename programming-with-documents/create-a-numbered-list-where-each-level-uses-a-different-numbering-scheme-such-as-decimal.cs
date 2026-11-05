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
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Create a new multilevel list based on the default list template.
        List list = doc.Lists.Add(ListTemplate.NumberDefault);

        // Level 1 – decimal numbering (1., 2., 3.).
        ListLevel level1 = list.ListLevels[0];
        level1.NumberFormat = "%1.";
        level1.NumberStyle = NumberStyle.Arabic;
        level1.Alignment = ListLevelAlignment.Left;

        // Level 2 – lower‑roman numbering (i., ii., iii.).
        ListLevel level2 = list.ListLevels[1];
        level2.NumberFormat = "%2.";
        level2.NumberStyle = NumberStyle.LowercaseRoman;
        level2.Alignment = ListLevelAlignment.Left;

        // Apply the list to the builder.
        builder.ListFormat.List = list;

        // First‑level item.
        builder.Writeln("First level item 1");

        // Second‑level items.
        builder.ListFormat.ListIndent();
        builder.Writeln("Second level item i");
        builder.Writeln("Second level item ii");
        builder.ListFormat.ListOutdent();

        // Another first‑level item.
        builder.Writeln("First level item 2");

        // End list formatting.
        builder.ListFormat.RemoveNumbers();

        // Save the document.
        string outputPath = "NumberedList.docx";
        doc.Save(outputPath);

        // Verify that the file was saved and can be reopened.
        if (File.Exists(outputPath))
        {
            Document loaded = new Document(outputPath);
            // Document successfully loaded – no further action required.
        }
    }
}
