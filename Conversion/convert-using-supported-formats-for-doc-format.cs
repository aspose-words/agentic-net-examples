using System;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Load an existing document. Aspose.Words automatically detects the format.
        Document doc = new Document("Input.docx");

        // Convert and save the document to the legacy DOC format using the SaveFormat enum.
        doc.Save("Output.doc", SaveFormat.Doc);

        // The same conversion can be performed with explicit save options.
        DocSaveOptions options = new DocSaveOptions(SaveFormat.Doc);
        doc.Save("OutputWithOptions.doc", options);
    }
}
