using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Verify that the document currently has no watermark.
        if (doc.Watermark.Type == WatermarkType.None)
        {
            // Since no watermark is present, add a text watermark.
            doc.Watermark.SetText("Confidential");
        }

        // Prepare the output folder and file path.
        string outputFolder = "Output";
        Directory.CreateDirectory(outputFolder);
        string outputFile = Path.Combine(outputFolder, "DocumentWithWatermark.docx");

        // Save the document with the newly added watermark.
        doc.Save(outputFile);
    }
}
