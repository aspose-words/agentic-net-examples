using System;
using Aspose.Words;
using Aspose.Words.Saving;

class RemoveWatermarkFromHtml
{
    static void Main()
    {
        // Path to the source HTML document.
        string inputPath = @"C:\Docs\SourceDocument.html";

        // Path where the resulting HTML document will be saved.
        string outputPath = @"C:\Docs\ResultDocument.html";

        // Load the HTML document.
        Document doc = new Document(inputPath);

        // If a watermark is present, remove it.
        if (doc.Watermark.Type != WatermarkType.None)
        {
            doc.Watermark.Remove();
        }

        // Save the document back to HTML format.
        // The SaveFormat is inferred from the file extension, but can be specified explicitly.
        doc.Save(outputPath, SaveFormat.Html);
    }
}
