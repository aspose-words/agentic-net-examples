using System;
using Aspose.Words;
using Aspose.Words.Lists;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Create a numbered list using a built‑in template.
        List list = doc.Lists.Add(ListTemplate.NumberDefault);

        // Apply the list to the builder.
        builder.ListFormat.List = list;

        // Write first‑level list items.
        builder.ListFormat.ListLevelNumber = 0;
        builder.Writeln("First item");
        builder.Writeln("Second item");

        // Write a nested second‑level list.
        builder.ListFormat.ListLevelNumber = 1;
        builder.Writeln("Nested item 1");
        builder.Writeln("Nested item 2");

        // Return to first level and add another item.
        builder.ListFormat.ListLevelNumber = 0;
        builder.Writeln("Third item");

        // Remove list formatting to avoid affecting subsequent paragraphs.
        builder.ListFormat.RemoveNumbers();

        // Save the document with compliance higher than Ecma376 to retain custom list settings.
        OoxmlSaveOptions saveOptions = new OoxmlSaveOptions(SaveFormat.Docx);
        saveOptions.Compliance = OoxmlCompliance.Iso29500_2008_Strict; // higher than Ecma376
        doc.Save("CustomList.docx", saveOptions);
    }
}
