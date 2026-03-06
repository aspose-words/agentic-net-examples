using System;
using Aspose.Words;
using Aspose.Words.Saving;

class ConvertToDocm
{
    static void Main()
    {
        // Load an existing document (DOCX, PDF, etc.).
        Document doc = new Document("InputDocument.docx");

        // Create OoxmlSaveOptions specifying the DOCM format.
        // This options class allows additional configuration for OOXML formats.
        OoxmlSaveOptions saveOptions = new OoxmlSaveOptions(SaveFormat.Docm);

        // Example of additional configuration (optional):
        // - Enable pretty formatting of the XML.
        // - Set a password to protect the macro-enabled document.
        saveOptions.PrettyFormat = true;
        saveOptions.Password = "SecurePassword123";

        // Save the document as a macro‑enabled DOCM file using the configured options.
        doc.Save("OutputDocument.docm", saveOptions);
    }
}
