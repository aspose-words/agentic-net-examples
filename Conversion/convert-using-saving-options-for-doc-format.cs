using System;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Load the source document (any supported format, e.g., DOCX).
        Document doc = new Document("Input.docx");

        // Create a DocSaveOptions instance for the older DOC format.
        DocSaveOptions saveOptions = new DocSaveOptions(SaveFormat.Doc);

        // Optional: set a password that will protect the saved DOC file.
        saveOptions.Password = "MyPassword";

        // Optional: preserve the routing slip if the document contains one.
        saveOptions.SaveRoutingSlip = true;

        // Save the document as a DOC file using the specified save options.
        doc.Save("Output.doc", saveOptions);
    }
}
