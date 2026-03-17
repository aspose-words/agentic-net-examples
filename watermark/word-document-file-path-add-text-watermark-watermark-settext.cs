using System;
using Aspose.Words;
using Aspose.Words.Drawing;

class WatermarkExample
{
    static void Main()
    {
        // Path to the existing Word document.
        string inputPath = @"C:\Docs\Sample.docx";

        // Load the document from the file system.
        Document doc = new Document(inputPath);

        // Add a plain text watermark to every page of the document.
        // The watermark text must be between 1 and 200 characters.
        doc.Watermark.SetText("Confidential");

        // Save the modified document to a new file (or overwrite the original).
        string outputPath = @"C:\Docs\Sample_With_Watermark.docx";
        doc.Save(outputPath);
    }
}
