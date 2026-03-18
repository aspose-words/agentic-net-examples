using Aspose.Words;
using Aspose.Words.Fields;

class Program
{
    static void Main()
    {
        // Load the document that contains the TOC field.
        Document doc = new Document("TOCDocument.docx");

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
            // The Field.Remove method removes the field from the document.
            tocField.Remove();
        }

        // Save the document after the TOC field has been removed.
        doc.Save("TOCDocument_NoTOC.docx");
    }
}
