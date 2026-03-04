using System;
using Aspose.Words;
using Aspose.Words.Saving;

class OdtToEpubConverter
{
    static void Main()
    {
        // Path to the source ODT file.
        string odtPath = @"C:\Input\source.odt";

        // Path where the resulting EPUB will be saved.
        string epubPath = @"C:\Output\result.epub";

        // Load the ODT document.
        Document doc = new Document(odtPath);

        // Configure EPUB save options.
        HtmlSaveOptions epubOptions = new HtmlSaveOptions(SaveFormat.Epub)
        {
            // Populate navigation map with headings up to level 3 (default).
            NavigationMapLevel = 3
        };

        // Save the document as EPUB with navigation support.
        doc.Save(epubPath, epubOptions);
    }
}
