using Aspose.Words;
using Aspose.Words.Lists;
using Aspose.Words;

class Program
{
    static void Main()
    {
        // Create a new document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add a numbered list and enable restart at each section.
        // This causes Aspose.Words to write the <w:restartNumber> tag when saved as DOC.
        List list = doc.Lists.Add(ListTemplate.NumberDefault);
        list.IsRestartAtEachSection = true;

        // Apply the list to the builder.
        builder.ListFormat.List = list;

        // First section – two list items.
        builder.Writeln("Item 1");
        builder.Writeln("Item 2");

        // Insert a section break to start a new section.
        builder.InsertBreak(BreakType.SectionBreakNewPage);

        // Second section – numbering restarts from 1.
        builder.Writeln("Item 1");
        builder.Writeln("Item 2");

        // Save the document in DOC format, which will contain the restartNum tag.
        doc.Save("RestartNum.doc");
    }
}
