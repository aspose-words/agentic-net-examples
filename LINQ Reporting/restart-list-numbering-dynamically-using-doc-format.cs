using System;
using Aspose.Words;
using Aspose.Words.Lists;

class RestartListNumbering
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Add a numbered list to the document's list collection using the built‑in NumberDefault template.
        List list = doc.Lists.Add(ListTemplate.NumberDefault);

        // Enable restarting the list numbering at each section (supported for DOC, RTF and DOCX).
        list.IsRestartAtEachSection = true;

        // Use DocumentBuilder to write paragraphs and apply the list.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.ListFormat.List = list;

        // First section items.
        builder.Writeln("Item 1 in first section");
        builder.Writeln("Item 2 in first section");
        builder.Writeln("Item 3 in first section");

        // Insert a section break (new page) – the list will restart after this.
        builder.InsertBreak(BreakType.SectionBreakNewPage);

        // Second section items – numbering restarts from 1.
        builder.Writeln("Item 1 in second section");
        builder.Writeln("Item 2 in second section");
        builder.Writeln("Item 3 in second section");

        // Optional: remove list formatting from the builder.
        builder.ListFormat.RemoveNumbers();

        // Save the document in DOC format to preserve the restart behavior.
        string outputPath = "RestartListNumbering.doc";
        doc.Save(outputPath, SaveFormat.Doc);
    }
}
