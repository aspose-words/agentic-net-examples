using System;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Load an existing document (any supported format, e.g., PDF, DOC, etc.).
        Document doc = new Document("input.pdf");

        // Save the document as DOCX using the SaveFormat enumeration.
        doc.Save("output.docx", SaveFormat.Docx);

        // Optionally, use OoxmlSaveOptions for additional control (e.g., compression level).
        OoxmlSaveOptions saveOptions = new OoxmlSaveOptions(SaveFormat.Docx);
        // saveOptions.CompressionLevel = CompressionLevel.Maximum; // Uncomment to set compression.
        doc.Save("output_with_options.docx", saveOptions);
    }
}
