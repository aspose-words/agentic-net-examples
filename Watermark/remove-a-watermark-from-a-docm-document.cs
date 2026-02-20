using System;
using Aspose.Words;
using Aspose.Words.Saving;

class RemoveWatermarkExample
{
    static void Main()
    {
        // Load the DOCM document.
        Document doc = new Document("InputDocument.docm");

        // If a watermark is present (text or image), remove it.
        if (doc.Watermark.Type != WatermarkType.None)
        {
            doc.Watermark.Remove();
        }

        // Save the document back to DOCM format.
        OoxmlSaveOptions saveOptions = new OoxmlSaveOptions(SaveFormat.Docm);
        doc.Save("OutputDocument.docm", saveOptions);
    }
}
