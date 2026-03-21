using Aspose.Words;
using Aspose.Words.Fields;

class Program
{
    static void Main()
    {
        // Create a new document and insert a TOC field.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.InsertTableOfContents("\\o \"1-3\" \\h \\z \\u");

        // Locate the first TOC field in the document.
        Field tocField = null;
        foreach (Field field in doc.Range.Fields)
        {
            if (field.Type == FieldType.FieldTOC)
            {
                tocField = field;
                break;
            }
        }

        // Remove the TOC field if it was found.
        if (tocField != null)
        {
            tocField.Remove();
        }

        // Save the document after the TOC field has been removed.
        doc.Save("TOCDocument_NoTOC.docx");
    }
}
