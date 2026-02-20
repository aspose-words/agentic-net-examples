using System;
using Aspose.Words;
using Aspose.Words.Loading;
using Aspose.Words.Saving;

class PdfToDotConverter
{
    static void Main()
    {
        // Load the PDF document using PdfLoadOptions.
        var loadOptions = new PdfLoadOptions();
        Document doc = new Document("input.pdf", loadOptions);

        // Prepare save options for the DOT (Word template) format.
        var saveOptions = new DocSaveOptions(SaveFormat.Dot);

        // Save the document as a DOT file.
        doc.Save("output.dot", saveOptions);
    }
}
