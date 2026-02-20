using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Configure save options to use a compliance level that supports non‑primitive shapes
        // such as group shapes (requires DML). ISO 29500 Transitional is sufficient.
        OoxmlSaveOptions saveOptions = new OoxmlSaveOptions(SaveFormat.Docx);
        saveOptions.Compliance = OoxmlCompliance.Iso29500_2008_Transitional;

        // Save the document as DOCX, ready for inserting a GroupShape later.
        doc.Save("GroupShapeReady.docx", saveOptions);
    }
}
