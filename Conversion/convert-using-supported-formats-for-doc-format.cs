using System;
using Aspose.Words;
using Aspose.Words.Saving;

namespace AsposeWordsExamples
{
    class ConvertToDoc
    {
        static void Main()
        {
            // Load an existing document (any supported format, e.g., DOCX).
            Document doc = new Document("InputDocument.docx");

            // Option 1: Save directly using the SaveFormat enumeration.
            doc.Save("OutputDocument.doc", SaveFormat.Doc);

            // Option 2: Use DocSaveOptions for additional control (e.g., password protection).
            DocSaveOptions options = new DocSaveOptions(SaveFormat.Doc);
            options.Password = "MyPassword";               // Optional: set a password.
            options.SaveRoutingSlip = true;                // Optional: preserve routing slip if present.

            doc.Save("OutputDocument_WithOptions.doc", options);
        }
    }
}
