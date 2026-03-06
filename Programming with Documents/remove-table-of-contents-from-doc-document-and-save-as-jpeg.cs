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

        // Remove every Table of Contents field from the document.
        // Iterate backwards because removing a field modifies the collection.
        for (int i = doc.Range.Fields.Count - 1; i >= 0; i--)
        {
            Field field = doc.Range.Fields[i];
            if (field.Type == FieldType.FieldTOC)
                field.Remove();
        }

        // Prepare JPEG image save options.
        ImageSaveOptions jpegOptions = new ImageSaveOptions(SaveFormat.Jpeg);
        // By default only the first page is rendered for image formats.
        // If you need a specific page, set jpegOptions.PageSet = new PageSet(pageIndex);

        // Save the modified document as a JPEG image.
        doc.Save("Output.jpg", jpegOptions);
    }
}
