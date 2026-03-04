using Aspose.Words;
using Aspose.Words.Lists;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Create a new document and a builder.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add a numbered list and enable restarting at each section.
        List list = doc.Lists.Add(ListTemplate.NumberDefault);
        list.IsRestartAtEachSection = true;

        // Apply the list to the first section.
        builder.ListFormat.List = list;
        builder.Writeln("Item 1");
        builder.Writeln("Item 2");

        // Insert a section break; the list will restart automatically.
        builder.InsertBreak(BreakType.SectionBreakNewPage);
        builder.Writeln("Item 1");
        builder.Writeln("Item 2");

        // End list formatting.
        builder.ListFormat.RemoveNumbers();

        // Save as a DOCM file with a compliance level that supports the property.
        OoxmlSaveOptions saveOptions = new OoxmlSaveOptions
        {
            Compliance = OoxmlCompliance.Iso29500_2008_Transitional,
            SaveFormat = SaveFormat.Docm
        };
        doc.Save("RestartList.docm", saveOptions);
    }
}
