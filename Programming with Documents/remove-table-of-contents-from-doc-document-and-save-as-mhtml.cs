using System;
using Aspose.Words;
using Aspose.Words.Fields;
using Aspose.Words.Saving;

class RemoveTocAndSaveAsMhtml
{
    static void Main()
    {
        // Load the existing DOC document.
        Document doc = new Document("InputDocument.doc");

        // Remove all Table of Contents (TOC) fields from the document.
        // Iterate backwards because removing a field changes the collection indexes.
        for (int i = doc.Range.Fields.Count - 1; i >= 0; i--)
        {
            Field field = doc.Range.Fields[i];
            if (field.Type == FieldType.FieldTOC)
                field.Remove();
        }

        // Prepare save options for MHTML format.
        HtmlSaveOptions saveOptions = new HtmlSaveOptions(SaveFormat.Mhtml);

        // Save the modified document as MHTML.
        doc.Save("OutputDocument.mhtml", saveOptions);
    }
}
