using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Fields;

class Program
{
    static void Main()
    {
        // Load the source DOC/DOCX document.
        Document doc = new Document("Input.docx");

        // Collect all Table of Contents (TOC) fields.
        List<Field> tocFields = new List<Field>();
        foreach (Field field in doc.Range.Fields)
        {
            if (field.Type == FieldType.FieldTOC)
                tocFields.Add(field);
        }

        // Remove each TOC field from the document.
        foreach (Field toc in tocFields)
        {
            toc.Remove(); // Removes the field start, separator, end and its result.
        }

        // Save the document as an EPUB file.
        doc.Save("Output.epub", SaveFormat.Epub);
    }
}
