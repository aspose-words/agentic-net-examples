using System;
using Aspose.Words;
using Aspose.Words.Lists;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Use DocumentBuilder to add content.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Create a multilevel list based on the default numbered template.
        List list = doc.Lists.Add(ListTemplate.NumberDefault);

        // Level 0 – decimal numbers (1., 2., 3., ...). This is the default, but set explicitly for clarity.
        list.ListLevels[0].NumberStyle = NumberStyle.Arabic;

        // Level 1 – lower‑roman numerals (i., ii., iii., ...).
        list.ListLevels[1].NumberStyle = NumberStyle.LowercaseRoman;

        // Apply the list to the builder.
        builder.ListFormat.List = list;

        // First top‑level item.
        builder.Writeln("First top‑level item");

        // Indent to level 1 (lower‑roman).
        builder.ListFormat.ListIndent();
        builder.Writeln("First sub‑item (lower‑roman)");

        // Another sub‑item.
        builder.Writeln("Second sub‑item (lower‑roman)");

        // Outdent back to top level.
        builder.ListFormat.ListOutdent();
        builder.Writeln("Second top‑level item");

        // Clean up list formatting.
        builder.ListFormat.RemoveNumbers();

        // Save the document to a file in the local folder.
        doc.Save("NumberedListDifferentLevels.docx");
    }
}
