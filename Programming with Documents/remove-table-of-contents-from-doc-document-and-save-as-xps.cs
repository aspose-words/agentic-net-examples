using System;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Words.Fields; // Added namespace for Field and FieldType

class RemoveTocAndSaveAsXps
{
    static void Main()
    {
        // Load the existing DOCX document.
        Document doc = new Document("input.docx");

        // Collect all TOC fields first to avoid modifying the collection while iterating.
        var tocFields = new System.Collections.Generic.List<Field>();
        foreach (Field field in doc.Range.Fields)
        {
            if (field.Type == FieldType.FieldTOC)
                tocFields.Add(field);
        }

        // Remove each TOC field.
        foreach (Field toc in tocFields)
        {
            toc.Remove();
        }

        // Update remaining fields (optional but recommended).
        doc.UpdateFields();

        // Save the document as XPS.
        XpsSaveOptions xpsOptions = new XpsSaveOptions();
        doc.Save("output.xps", xpsOptions);
    }
}
