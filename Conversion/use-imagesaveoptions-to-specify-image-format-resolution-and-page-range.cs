using System;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Load the source DOCX document.
        Document doc = new Document("Input.docx");

        // Create ImageSaveOptions for PNG output.
        ImageSaveOptions options = new ImageSaveOptions(SaveFormat.Png);

        // Set the desired resolution (dots per inch).
        options.Resolution = 300;

        // Define the page range to render (pages 2‑4, zero‑based indices).
        // PageSet(startPage, endPage) includes both start and end pages.
        options.PageSet = new PageSet(1, 3);

        // Save the selected pages as a PNG image.
        doc.Save("Output_Page2to4.png", options);
    }
}
