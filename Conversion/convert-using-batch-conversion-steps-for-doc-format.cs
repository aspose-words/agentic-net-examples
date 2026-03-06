using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

class BatchConversion
{
    static void Main()
    {
        // Load the source document. Aspose.Words automatically detects the format.
        Document doc = new Document("input.docx");

        // Create save options for the legacy DOC format.
        DocSaveOptions saveOptions = new DocSaveOptions(SaveFormat.Doc);

        // Optionally set additional options, e.g., a password.
        // saveOptions.Password = "MyPassword";

        // Save the document in DOC format using the specified options.
        doc.Save("output.doc", saveOptions);
    }
}
