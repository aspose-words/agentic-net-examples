using System;
using Aspose.Words;
using Aspose.Words.Lists;

public class RestartListNumbering
{
    public static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add a list based on the default numbered template.
        List list = doc.Lists.Add(ListTemplate.NumberDefault);

        // First section – use the list as is.
        list.ListLevels[0].StartAt = 1; // Ensure numbering starts at 1.
        builder.ListFormat.List = list;
        builder.Writeln("Section 1 – Item 1");
        builder.Writeln("Section 1 – Item 2");
        builder.ListFormat.RemoveNumbers();

        // Insert a section break (new page) to start a new section.
        builder.InsertBreak(BreakType.SectionBreakNewPage);

        // Second section – reset the starting number before re‑using the same list.
        list.ListLevels[0].StartAt = 1; // Reset numbering for the new section.
        builder.ListFormat.List = list;
        builder.Writeln("Section 2 – Item 1");
        builder.Writeln("Section 2 – Item 2");
        builder.ListFormat.RemoveNumbers();

        // Save the document to disk.
        doc.Save("RestartNumbering.docx");
    }
}
