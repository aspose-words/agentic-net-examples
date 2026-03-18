using System.Drawing;
using Aspose.Words;
using Aspose.Words.Lists;

class Program
{
    static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Create a list based on a predefined template.
        List list = doc.Lists.Add(ListTemplate.NumberDefault);
        // Ensure the list starts at 1.
        list.ListLevels[0].StartAt = 1;

        // ---------- First section ----------
        builder.Writeln("Section 1:");
        // Apply the list to the paragraphs in this section.
        builder.ListFormat.List = list;
        builder.Writeln("Item 1");
        builder.Writeln("Item 2");
        builder.Writeln("Item 3");
        // End the list for this section.
        builder.ListFormat.RemoveNumbers();

        // Insert a section break to start a new section.
        builder.InsertBreak(BreakType.SectionBreakNewPage);

        // Reset the starting number before using the list in the new section.
        list.ListLevels[0].StartAt = 1;

        // ---------- Second section ----------
        builder.Writeln("Section 2:");
        // Apply the same list again; numbering will restart from 1.
        builder.ListFormat.List = list;
        builder.Writeln("Item 1");
        builder.Writeln("Item 2");
        builder.Writeln("Item 3");
        builder.Writeln("Item 4");
        // End the list for this section.
        builder.ListFormat.RemoveNumbers();

        // Save the document to disk.
        doc.Save("RestartNumberingEachSection.docx");
    }
}
