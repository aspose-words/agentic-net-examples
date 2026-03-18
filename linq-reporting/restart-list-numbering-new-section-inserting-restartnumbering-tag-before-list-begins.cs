using System;
using Aspose.Words;
using Aspose.Words.Lists;
using Aspose.Words.Saving;

class RestartListNumbering
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Initialize a DocumentBuilder for editing the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add a default numbered list to the document's list collection.
        List list = doc.Lists.Add(ListTemplate.NumberDefault);

        // Enable restarting of the list numbering at each new section.
        list.IsRestartAtEachSection = true;

        // Apply the list to the following paragraphs.
        builder.ListFormat.List = list;

        // First section items.
        builder.Writeln("Item 1");
        builder.Writeln("Item 2");

        // Insert a section break – the list will restart after this break.
        builder.InsertBreak(BreakType.SectionBreakNewPage);

        // Second section items – numbering starts again from 1.
        builder.Writeln("Item 1");
        builder.Writeln("Item 2");

        // Save the document with a compliance level that supports the restart flag.
        OoxmlSaveOptions saveOptions = new OoxmlSaveOptions
        {
            Compliance = OoxmlCompliance.Iso29500_2008_Transitional
        };
        doc.Save("RestartList.docx", saveOptions);
    }
}
