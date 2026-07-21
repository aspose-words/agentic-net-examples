using System;
using Aspose.Words;
using Aspose.Words.Lists;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Create a multilevel list based on the default numbered template.
        List list = doc.Lists.Add(ListTemplate.NumberDefault);

        // Level 0 – decimal numbers (1., 2., 3., ...).
        ListLevel level0 = list.ListLevels[0];
        level0.NumberStyle = NumberStyle.Arabic;

        // Level 1 – lower‑roman numbers (i., ii., iii., ...).
        ListLevel level1 = list.ListLevels[1];
        level1.NumberStyle = NumberStyle.LowercaseRoman;

        // Use a DocumentBuilder to add paragraphs that use the list.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.ListFormat.List = list;

        // First top‑level item.
        builder.Writeln("First item (decimal)");

        // Indent to level 1 – will use lower‑roman numbering.
        builder.ListFormat.ListIndent();
        builder.Writeln("Sub‑item (lower‑roman)");
        builder.Writeln("Another sub‑item (lower‑roman)");

        // Outdent back to level 0.
        builder.ListFormat.ListOutdent();
        builder.Writeln("Second item (decimal)");

        // Remove list formatting from subsequent paragraphs.
        builder.ListFormat.RemoveNumbers();

        // Save the document to disk.
        doc.Save("NumberedList.docx");
    }
}
