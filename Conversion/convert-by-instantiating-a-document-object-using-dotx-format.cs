using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Load an existing document (any supported format, e.g., DOCX)
        Document doc = new Document("input.docx");

        // Set up save options to specify DOTX (macro‑free template) format
        OoxmlSaveOptions saveOptions = new OoxmlSaveOptions(SaveFormat.Dotx);

        // Save the document as a DOTX file
        doc.Save("output.dotx", saveOptions);
    }
}
