using System;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Words.Fields;

class Program
{
    static void Main()
    {
        // Load the source DOC document.
        Document doc = new Document("Input.doc");

        // Remove every Table of Contents (TOC) field from the document.
        foreach (Field field in doc.Range.Fields)
        {
            if (field.Type == FieldType.FieldTOC)
                field.Remove();
        }

        // Save the modified document as a PNG image (first page only).
        ImageSaveOptions pngOptions = new ImageSaveOptions(SaveFormat.Png)
        {
            // Render only the first page (page index is zero‑based).
            PageSet = new PageSet(0)
        };
        doc.Save("Output.png", pngOptions);
    }
}
