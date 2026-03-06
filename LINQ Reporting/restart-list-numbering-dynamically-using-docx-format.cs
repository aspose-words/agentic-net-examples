using Aspose.Words;
using Aspose.Words.Lists;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Create a new document and a builder to add content.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add a numbered list definition to the document.
        List list = doc.Lists.Add(ListTemplate.NumberDefault);

        // Enable restarting the list numbering at each new section.
        // This flag is only written to DOCX when the OOXML compliance level is higher than Ecma376_2006.
        list.IsRestartAtEachSection = true;

        // Apply the list to the first set of paragraphs.
        builder.ListFormat.List = list;
        builder.Writeln("Item 1");
        builder.Writeln("Item 2");

        // Insert a section break; the list will restart numbering in the next section.
        builder.InsertBreak(BreakType.SectionBreakNewPage);
        builder.Writeln("Item 1 (new section)");
        builder.Writeln("Item 2 (new section)");

        // Optionally stop list formatting for any following paragraphs.
        builder.ListFormat.RemoveNumbers();

        // Save the document with a compliance level that supports the restart flag.
        OoxmlSaveOptions saveOptions = new OoxmlSaveOptions
        {
            Compliance = OoxmlCompliance.Iso29500_2008_Transitional
        };
        doc.Save("RestartList.docx", saveOptions);
    }
}
