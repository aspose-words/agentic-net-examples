using System;
using Aspose.Words;

class RemoveWatermarksExample
{
    static void Main()
    {
        // Path to the folder that contains the input document.
        // Replace with your actual folder path.
        string docsPath = @"C:\Docs\";

        // Load the existing Word document.
        Document doc = new Document(docsPath + "InputDocument.docx");

        // If a watermark is present, remove it.
        if (doc.Watermark.Type != WatermarkType.None)
        {
            doc.Watermark.Remove();
        }

        // Save the document without the watermark.
        doc.Save(docsPath + "OutputDocument.docx");
    }
}
