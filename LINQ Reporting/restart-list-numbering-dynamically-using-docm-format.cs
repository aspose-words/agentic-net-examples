using System;
using Aspose.Words;
using Aspose.Words.Lists;
using Aspose.Words.Saving;

class RestartListNumbering
{
    static void Main()
    {
        // Create a new document and a builder.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add a numbered list to the document's list collection.
        List list = doc.Lists.Add(ListTemplate.NumberDefault);

        // Enable restarting the list at each section.
        list.IsRestartAtEachSection = true;

        // Apply the list to the first section.
        builder.ListFormat.List = list;
        builder.Writeln("Item 1");
        builder.Writeln("Item 2");

        // Insert a section break (new page) – the list will restart after this break.
        builder.InsertBreak(BreakType.SectionBreakNewPage);

        // Continue using the same list; numbering restarts automatically.
        builder.Writeln("Item 1 (new section)");
        builder.Writeln("Item 2 (new section)");

        // Remove list formatting from subsequent paragraphs if needed.
        builder.ListFormat.RemoveNumbers();

        // Save the document as a macro‑enabled DOCM file.
        // The compliance level must be higher than Ecma376_2006 for the restart flag to be written.
        OoxmlSaveOptions saveOptions = new OoxmlSaveOptions
        {
            Compliance = OoxmlCompliance.Iso29500_2008_Transitional,
            SaveFormat = SaveFormat.Docm
        };

        doc.Save("RestartListNumbering.docm", saveOptions);
    }
}
