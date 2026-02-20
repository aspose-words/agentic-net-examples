using System;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Load an existing document (any supported format).
        Document doc = new Document("input.docx");

        // Create save options for the legacy DOC format.
        DocSaveOptions saveOptions = new DocSaveOptions();

        // Example option: include Aspose.Words generator name in the output.
        saveOptions.ExportGeneratorName = true;

        // Example option: set a password for the DOC file (optional).
        // saveOptions.Password = "myPassword";

        // Save the document as DOC using the specified options.
        doc.Save("output.doc", saveOptions);
    }
}
