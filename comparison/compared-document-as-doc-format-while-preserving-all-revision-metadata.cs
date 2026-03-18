using System;
using Aspose.Words;
using Aspose.Words.Saving;

namespace AsposeWordsRevisionSave
{
    class Program
    {
        static void Main()
        {
            // Load the original and edited documents.
            Document original = new Document("Original.docx");
            Document edited = new Document("Edited.docx");

            // Compare the documents. The revisions (track changes) will be added to the original document.
            original.Compare(edited, "Comparer", DateTime.Now);

            // Prepare save options for the legacy DOC format.
            DocSaveOptions saveOptions = new DocSaveOptions(SaveFormat.Doc);

            // Save the document with revisions preserved.
            original.Save("Compared.doc", saveOptions);
        }
    }
}
