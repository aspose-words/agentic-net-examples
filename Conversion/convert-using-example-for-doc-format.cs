using System;
using Aspose.Words;

namespace AsposeWordsExamples
{
    class ConvertToDocExample
    {
        static void Main()
        {
            // Path to the folder containing input and output files.
            string dataDir = @"C:\Data\";

            // Load an existing DOCX document from the file system.
            Document doc = new Document(dataDir + "Document.docx");

            // Save the document in the older Microsoft Word 97‑2007 DOC format.
            doc.Save(dataDir + "Document.ConvertToDoc.doc", SaveFormat.Doc);
        }
    }
}
