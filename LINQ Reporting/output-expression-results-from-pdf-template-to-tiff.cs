using System;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Load the PDF template (the document may contain fields that need to be evaluated)
        Document doc = new Document("Template.pdf");

        // Update all fields (e.g., MERGEFIELD, FORMFIELD) so that their expression results are reflected
        doc.UpdateFields();

        // Configure image save options for TIFF output.
        // Use the TiffFrames layout so each page becomes a separate frame in a multi‑frame TIFF.
        ImageSaveOptions tiffOptions = new ImageSaveOptions(SaveFormat.Tiff)
        {
            PageLayout = MultiPageLayout.TiffFrames()
        };

        // Save the document as a multi‑frame TIFF image.
        doc.Save("Result.tiff", tiffOptions);
    }
}
