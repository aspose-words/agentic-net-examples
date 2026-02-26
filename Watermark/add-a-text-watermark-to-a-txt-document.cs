using System;
using Aspose.Words;

class Program
{
    static void Main()
    {
        // Create a new Word document (it can be saved as TXT later)
        Document doc = new Document();

        // Add a text watermark to the document
        doc.Watermark.SetText("Confidential");

        // Save the document as a plain‑text file. Note that the watermark will not be visible in the TXT output
        // because plain‑text format does not support watermarks.
        doc.Save("WatermarkedDocument.txt");
    }
}
