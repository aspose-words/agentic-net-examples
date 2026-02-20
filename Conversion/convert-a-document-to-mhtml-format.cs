using System;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Load the source document (replace with your actual file path).
        Document doc = new Document("input.docx");

        // Configure save options to produce MHTML output.
        HtmlSaveOptions saveOptions = new HtmlSaveOptions(SaveFormat.Mhtml);

        // Save the document as an MHTML file.
        doc.Save("output.mht", saveOptions);
    }
}
