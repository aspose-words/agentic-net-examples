using System;
using Aspose.Words;
using Aspose.Words.Saving;

class DocmToEpubConverter
{
    static void Main()
    {
        // Load the macro-enabled DOCM document from disk.
        Document doc = new Document(@"C:\Input\sample.docm");

        // Save the document as EPUB. The file extension determines the format,
        // but we can also explicitly set the SaveFormat via HtmlSaveOptions.
        HtmlSaveOptions saveOptions = new HtmlSaveOptions(SaveFormat.Epub);
        doc.Save(@"C:\Output\sample.epub", saveOptions);
    }
}
