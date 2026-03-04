using System;
using Aspose.Words;
using Aspose.Words.Saving;
using System.Drawing;

class Program
{
    static void Main()
    {
        // Load an existing Word document.
        Document doc = new Document("Input.docx");

        // Set up options for rendering the document to an image.
        ImageSaveOptions options = new ImageSaveOptions(SaveFormat.Png);
        options.Resolution = 300;               // Render at 300 DPI.
        options.PageSet = new PageSet(0);       // Convert only the first page.
        options.PaperColor = Color.Transparent; // Optional: transparent background.

        // Save the rendered page as an image file.
        doc.Save("Output.png", options);
    }
}
