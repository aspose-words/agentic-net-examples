using System;
using Aspose.Words;

namespace AsposeWordsConversion
{
    class Program
    {
        static void Main()
        {
            // Load an existing document (any supported format, e.g., DOCX).
            Document doc = new Document("InputDocument.docx");

            // Save the document in the legacy Microsoft Word 97‑2007 DOC format.
            doc.Save("OutputDocument.doc", SaveFormat.Doc);
        }
    }
}
