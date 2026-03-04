using System;
using Aspose.Words;
using Aspose.Words.Lists;

class InsertRestartNumTag
{
    static void Main()
    {
        // Create a new document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add a numbered list.
        List list = doc.Lists.Add(ListTemplate.NumberDefault);
        // Enable restarting the list numbering at each new section.
        list.IsRestartAtEachSection = true;

        // Apply the list to the builder.
        builder.ListFormat.List = list;

        // First section – first two list items.
        builder.Writeln("Item 1");
        builder.Writeln("Item 2");

        // Insert a section break to trigger a restart of the list numbering.
        builder.InsertBreak(BreakType.SectionBreakNewPage);

        // Second section – list numbering restarts from 1.
        builder.Writeln("Item 1 (restarted)");
        builder.Writeln("Item 2 (restarted)");

        // Save the document in DOCM format (macro-enabled Word document).
        doc.Save("RestartNumberedParagraphs.docm");
    }
}
