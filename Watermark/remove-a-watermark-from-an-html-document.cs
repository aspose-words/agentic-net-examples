using System;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Words.Loading;

class RemoveWatermarkFromHtml
{
    static void Main()
    {
        // Path to the source HTML file.
        string inputPath = @"C:\Docs\input.html";

        // Load the HTML document.
        // HtmlLoadOptions can be customized if needed; using defaults here.
        Document doc = new Document(inputPath, new HtmlLoadOptions());

        // If a watermark is present (text or image), remove it.
        if (doc.Watermark.Type != WatermarkType.None)
        {
            doc.Watermark.Remove();
        }

        // Path to the output HTML file.
        string outputPath = @"C:\Docs\output.html";

        // Save the document back to HTML format.
        // HtmlSaveOptions can be customized; using defaults here.
        doc.Save(outputPath, new HtmlSaveOptions());

        Console.WriteLine("Watermark removal completed.");
    }
}
