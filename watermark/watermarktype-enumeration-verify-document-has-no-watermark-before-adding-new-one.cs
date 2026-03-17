using System;
using Aspose.Words;
using Aspose.Words.Drawing;
using System.Drawing;

class WatermarkExample
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Verify that the document does not already contain a watermark.
        // Watermark.Type returns WatermarkType.None when no watermark is set.
        if (doc.Watermark.Type == WatermarkType.None)
        {
            // Since there is no existing watermark, add a new text watermark.
            // You can customize the appearance with TextWatermarkOptions if needed.
            doc.Watermark.SetText("Confidential");
        }
        else
        {
            // If a watermark already exists, you could handle it here (e.g., remove or replace).
            // For this example we simply leave the existing watermark unchanged.
            Console.WriteLine("Document already contains a watermark of type: " + doc.Watermark.Type);
        }

        // Save the document to a file.
        doc.Save("Result.docx");
    }
}
