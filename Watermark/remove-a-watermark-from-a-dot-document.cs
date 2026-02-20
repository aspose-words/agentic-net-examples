using System;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Load the existing DOT document.
        Document doc = new Document("InputTemplate.dot");

        // Remove the watermark if one is present.
        if (doc.Watermark.Type != WatermarkType.None)
            doc.Watermark.Remove();

        // Save the document back to DOT format.
        DocSaveOptions saveOptions = new DocSaveOptions(SaveFormat.Dot);
        doc.Save("OutputTemplate.dot", saveOptions);
    }
}
