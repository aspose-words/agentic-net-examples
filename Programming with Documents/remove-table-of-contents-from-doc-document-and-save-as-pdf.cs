using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Fields;

class Program
{
    static void Main()
    {
        // Load the source DOC document.
        Document doc = new Document("Input.doc");

        // Find all Table of Contents (TOC) fields in the document.
        List<Field> tocFields = new List<Field>();
        foreach (Field field in doc.Range.Fields)
        {
            if (field.Type == FieldType.FieldTOC)
                tocFields.Add(field);
        }

        // Remove each TOC field from the document.
        foreach (Field toc in tocFields)
        {
            toc.Remove();
        }

        // Save the resulting document as PDF.
        doc.Save("Output.pdf");
    }
}
