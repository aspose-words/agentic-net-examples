using System;
using Aspose.Words;
using Aspose.Words.Saving;

class ThumbnailGenerator
{
    static void Main()
    {
        // Load the source Word document that contains JPEG images.
        Document doc = new Document("Input.docx");

        // Configure image save options to render the document pages as JPEG thumbnails.
        // The Scale property reduces both width and height to 50% of the original size.
        ImageSaveOptions saveOptions = new ImageSaveOptions(SaveFormat.Jpeg)
        {
            Scale = 0.5f,            // 50% of original dimensions
            JpegQuality = 80        // Optional: set JPEG quality (0‑100)
        };

        // Save the rendered page (or the whole document) as a JPEG thumbnail.
        // If the document has multiple pages, each page will be saved separately
        // using the PageSet property in a loop (shown below).
        for (int i = 0; i < doc.PageCount; i++)
        {
            // Render only the current page.
            saveOptions.PageSet = new PageSet(i);

            // Save the thumbnail for the current page.
            string outputPath = $"Thumbnail_Page{i + 1}.jpg";
            doc.Save(outputPath, saveOptions);
        }
    }
}
