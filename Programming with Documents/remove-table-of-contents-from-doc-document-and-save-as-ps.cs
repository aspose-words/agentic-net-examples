using System;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Words.Fields; // Added for Field and FieldType

class Program
{
    static void Main()
    {
        // Load the existing DOC document.
        Document doc = new Document("Input.doc");

        // Iterate through all fields in the document and remove any Table of Contents fields.
        foreach (Field field in doc.Range.Fields)
        {
            if (field.Type == FieldType.FieldTOC)
                field.Remove();
        }

        // Configure PostScript save options.
        PsSaveOptions psOptions = new PsSaveOptions
        {
            SaveFormat = SaveFormat.Ps
        };

        // Save the modified document as a PostScript file.
        doc.Save("Output.ps", psOptions);
    }
}
