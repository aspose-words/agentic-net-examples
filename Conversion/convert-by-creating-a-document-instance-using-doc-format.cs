using System;
using Aspose.Words;

namespace AsposeWordsExample
{
    class Program
    {
        static void Main()
        {
            // Create a blank Word document.
            Document doc = new Document();

            // Add a simple paragraph so the file is not empty.
            doc.FirstSection.Body.FirstParagraph.AppendChild(new Run(doc, "Hello world!"));

            // Save the document in the legacy DOC format (Microsoft Word 97‑2007).
            doc.Save("Result.doc", SaveFormat.Doc);
        }
    }
}
