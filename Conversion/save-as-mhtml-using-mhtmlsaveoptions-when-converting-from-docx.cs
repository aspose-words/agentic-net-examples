using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

class ConvertDocxToMhtml
{
    static void Main()
    {
        // Path to the source DOCX file.
        string inputPath = Path.Combine(Environment.CurrentDirectory, "Document.docx");

        // Path where the MHTML file will be saved.
        string outputPath = Path.Combine(Environment.CurrentDirectory, "Document.mht");

        // Load the DOCX document.
        Document doc = new Document(inputPath);

        // Create save options for MHTML format.
        // The constructor takes a SaveFormat that can be Html, Mhtml, Epub, Azw3 or Mobi.
        HtmlSaveOptions saveOptions = new HtmlSaveOptions(SaveFormat.Mhtml);

        // Save the document as MHTML using the specified options.
        doc.Save(outputPath, saveOptions);
    }
}
