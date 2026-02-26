using System;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Load an existing document. The format is detected automatically (e.g., DOCX, PDF, etc.).
        Document doc = new Document("Input.docx");

        // Save the document in the legacy Microsoft Word 97‑2007 DOC format using the SaveFormat enum.
        doc.Save("Output.doc", SaveFormat.Doc);

        // Alternatively, use DocSaveOptions to specify additional save settings such as a password
        // and whether to preserve the routing slip when saving to the DOC format.
        DocSaveOptions saveOptions = new DocSaveOptions(SaveFormat.Doc);
        saveOptions.Password = "MyPassword";      // Protect the document with a password.
        saveOptions.SaveRoutingSlip = true;       // Preserve any routing slip data.

        // Save the document with the specified options.
        doc.Save("OutputWithOptions.doc", saveOptions);
    }
}
