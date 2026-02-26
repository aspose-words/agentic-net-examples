using Aspose.Words;
using Aspose.Words.Lists;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Load the existing DOCX document.
        Document doc = new Document("input.docx");

        // Set each list in the document to restart numbering at the start of every section.
        foreach (List list in doc.Lists)
        {
            list.IsRestartAtEachSection = true;
        }

        // Save the document with a compliance level that supports the IsRestartAtEachSection property.
        OoxmlSaveOptions saveOptions = new OoxmlSaveOptions
        {
            Compliance = OoxmlCompliance.Iso29500_2008_Transitional
        };
        doc.Save("output.docx", saveOptions);
    }
}
