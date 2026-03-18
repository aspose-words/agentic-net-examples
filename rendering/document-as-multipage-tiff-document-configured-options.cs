using System;
using Aspose.Words;
using Aspose.Words.Saving;
using System.Drawing;

class Program
{
    static void Main()
    {
        // Create a simple document with three pages.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Page 1.");
        builder.InsertBreak(BreakType.PageBreak);
        builder.Writeln("Page 2.");
        builder.InsertBreak(BreakType.PageBreak);
        builder.Writeln("Page 3.");

        // Configure save options for a multipage TIFF.
        ImageSaveOptions tiffOptions = new ImageSaveOptions(SaveFormat.Tiff);
        // Use the TiffFrames layout so each page is saved as a separate frame.
        tiffOptions.PageLayout = MultiPageLayout.TiffFrames();
        // Optional: set resolution (dpi) for the rendered images.
        tiffOptions.Resolution = 300;

        // Save the document as a multipage TIFF file.
        doc.Save("MultipageOutput.tiff", tiffOptions);
    }
}
