using System;
using Aspose.Words;
using Aspose.Words.Saving;

class ListToEpubConverter
{
    static void Main()
    {
        // Path to the source DOC file that contains the list.
        string sourceDocPath = "input.doc";

        // Path where the resulting EPUB file will be saved.
        string outputEpubPath = "output.epub";

        // Load the source document.
        Document sourceDoc = new Document(sourceDocPath);

        // Ensure list labels are up‑to‑date so they appear in the plain‑text output.
        sourceDoc.UpdateListLabels();

        // Extract the entire document as plain text (list labels are included).
        string plainText = sourceDoc.ToString(SaveFormat.Text);

        // Create a new blank document and write the plain‑text content into it.
        Document epubDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(epubDoc);
        builder.Writeln(plainText);

        // Configure save options for EPUB output.
        HtmlSaveOptions epubOptions = new HtmlSaveOptions();
        epubOptions.SaveFormat = SaveFormat.Epub;   // Specify EPUB format.
        // Optional: keep default split criteria (single HTML part) and other settings.

        // Save the new document as an EPUB file.
        epubDoc.Save(outputEpubPath, epubOptions);
    }
}
