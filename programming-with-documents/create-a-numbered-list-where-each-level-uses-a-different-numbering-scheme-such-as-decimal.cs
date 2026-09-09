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

        // Create a DocumentBuilder to insert content.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add a multilevel list based on the default numbered template.
        List list = doc.Lists.Add(ListTemplate.NumberDefault);

        // Level 0 – Arabic numbers (1., 2., 3., ...). This is the default, but set explicitly.
        list.ListLevels[0].NumberStyle = NumberStyle.Arabic;

        // Level 1 – Lowercase Roman numerals (i., ii., iii., ...).
        list.ListLevels[1].NumberStyle = NumberStyle.LowercaseRoman;

        // Apply the list to the builder.
        builder.ListFormat.List = list;

        // First top‑level item.
        builder.Writeln("First top‑level item");

        // Indent to level 1 (lower‑roman).
        builder.ListFormat.ListIndent();
        builder.Writeln("First sub‑item");
        builder.Writeln("Second sub‑item");

        // Outdent back to level 0.
        builder.ListFormat.ListOutdent();
        builder.Writeln("Second top‑level item");

        // Another sub‑level example.
        builder.ListFormat.ListIndent();
        builder.Writeln("Another sub‑item");
        builder.ListFormat.ListOutdent();

        // Finish the list.
        builder.ListFormat.RemoveNumbers();

        // Ensure the output directory exists.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "NumberedList.docx");
        doc.Save(outputPath);
    }
}
