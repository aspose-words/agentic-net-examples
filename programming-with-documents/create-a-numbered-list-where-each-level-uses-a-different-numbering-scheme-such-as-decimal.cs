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

        // Create a multilevel list based on the default numbered template.
        List list = doc.Lists.Add(ListTemplate.NumberDefault);

        // Level 0 – decimal numbers (default, but set explicitly for clarity).
        list.ListLevels[0].NumberStyle = NumberStyle.Arabic;

        // Level 1 – lower‑roman numerals.
        list.ListLevels[1].NumberStyle = NumberStyle.LowercaseRoman;

        // Use DocumentBuilder to add paragraphs that use the list.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.ListFormat.List = list;

        // Top‑level items (decimal).
        builder.Writeln("Item 1");
        builder.Writeln("Item 2");

        // Indent to level 1 (lower‑roman).
        builder.ListFormat.ListIndent();
        builder.Writeln("Subitem i");
        builder.Writeln("Subitem ii");

        // Return to top level.
        builder.ListFormat.ListOutdent();

        // Finish the list.
        builder.ListFormat.RemoveNumbers();

        // Save the document to the current directory.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "NumberedList.docx");
        doc.Save(outputPath);
    }
}
