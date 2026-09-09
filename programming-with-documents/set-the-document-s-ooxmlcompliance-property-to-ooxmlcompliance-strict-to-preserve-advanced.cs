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

        // Add a numbered list to the document.
        doc.Lists.Add(ListTemplate.NumberDefault);
        List list = doc.Lists[0];

        // Enable restarting the list numbering at each new section.
        list.IsRestartAtEachSection = true;

        // Apply the list to the builder so that subsequent paragraphs become list items.
        builder.ListFormat.List = list;

        // Write some list items, insert a section break, then write more items.
        builder.Writeln("Item 1");
        builder.Writeln("Item 2");
        builder.InsertBreak(BreakType.SectionBreakNewPage);
        builder.Writeln("Item 1");
        builder.Writeln("Item 2");

        // Save the document with strict OOXML compliance to preserve advanced list settings.
        OoxmlSaveOptions saveOptions = new OoxmlSaveOptions(SaveFormat.Docx);
        saveOptions.Compliance = OoxmlCompliance.Iso29500_2008_Strict;

        doc.Save("AdvancedListSettings.docx", saveOptions);
    }
}
