using System;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Load the source Word document.
        Document doc = new Document("InputDocument.docx");

        // Create image save options for PNG format.
        ImageSaveOptions saveOptions = new ImageSaveOptions(SaveFormat.Png);

        // Render only a specific page (zero‑based index). 
        // For example, to render the second page set the index to 1.
        saveOptions.PageSet = new PageSet(1);

        // Save the selected page as a PNG image.
        doc.Save("Page2.png", saveOptions);
    }
}
