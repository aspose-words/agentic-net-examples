using System;
using Aspose.Words;

class Program
{
    static void Main()
    {
        // Paths to the source DOT file and the output file without watermark.
        string inputPath = @"C:\Docs\Template.dot";
        string outputPath = @"C:\Docs\Template_NoWatermark.dot";

        // Load the DOT document.
        Document doc = new Document(inputPath);

        // If a watermark exists, remove it.
        if (doc.Watermark.Type != WatermarkType.None)
        {
            doc.Watermark.Remove();
        }

        // Save the modified document.
        doc.Save(outputPath);
    }
}
