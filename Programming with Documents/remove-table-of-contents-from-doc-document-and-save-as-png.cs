using System;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Words.Fields; // Added for Field and FieldType

class Program
{
    static void Main()
    {
        // Load the source DOC document.
        Document doc = new Document("Input.doc");

        // Iterate through all fields in the document and remove any Table of Contents (TOC) fields.
        foreach (Field field in doc.Range.Fields)
        {
            if (field.Type == FieldType.FieldTOC)
                field.Remove();
        }

        // Save the modified document as a PNG image.
        // This renders the first page of the document to a PNG file.
        doc.Save("Output.png", SaveFormat.Png);
    }
}
