using System;
using Aspose.Words;
using Aspose.Words.Loading;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Load the macro‑enabled template (DOTM) from disk.
        // Using LoadOptions makes the format explicit and avoids any ambiguity.
        LoadOptions loadOptions = new LoadOptions { LoadFormat = LoadFormat.Dotm };
        Document document = new Document("input.dotm", loadOptions);

        // Create save options for EPUB output.
        // HtmlSaveOptions is used for HTML‑based formats, including EPUB.
        HtmlSaveOptions epubOptions = new HtmlSaveOptions(SaveFormat.Epub);

        // Save the document as an EPUB file.
        document.Save("output.epub", epubOptions);
    }
}
