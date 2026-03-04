using System;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Saving;

class ConvertDocumentToImages
{
    static void Main()
    {
        // Paths to the source document and the folder where images will be saved.
        string sourcePath = @"C:\Docs\MultiPage.docx";
        string outputFolder = @"C:\Output\";

        // Load the multi‑page Word document.
        Document doc = new Document(sourcePath);

        // Create ImageSaveOptions for JPEG output.
        ImageSaveOptions options = new ImageSaveOptions(SaveFormat.Jpeg);

        // Set custom rendering options (optional).
        options.Resolution = 300;                     // 300 DPI.
        options.ImageSize = new Size(2480, 3508);     // A4 size at 300 DPI.

        // Define the set of pages to render: pages 2 through 4 (zero‑based indices 1‑3).
        // This PageSet will be used to limit the pages we process.
        PageSet pagesToRender = new PageSet(new PageRange(1, 3));

        // Iterate over the desired page indices and save each page as a separate image.
        for (int pageIndex = 1; pageIndex <= 3; pageIndex++)
        {
            // Restrict the current save operation to a single page.
            options.PageSet = new PageSet(pageIndex);

            // Build the output file name, e.g., "Page_2.jpg".
            string outputPath = $"{outputFolder}Page_{pageIndex + 1}.jpg";

            // Save the selected page as an image.
            doc.Save(outputPath, options);
        }
    }
}
