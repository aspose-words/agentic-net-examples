using System;
using Aspose.Words;

class RemoveWatermarkFromDot
{
    static void Main()
    {
        // Path to the folder that contains the DOT template.
        string dataDir = @"C:\Docs\";

        // Load the DOT document.
        Document doc = new Document(dataDir + "Template.dot");

        // If a watermark is present, remove it.
        if (doc.Watermark.Type != WatermarkType.None)
        {
            doc.Watermark.Remove();
        }

        // Save the document without the watermark.
        doc.Save(dataDir + "Template_NoWatermark.dot");
    }
}
