using System;
using Aspose.Words;

namespace InsertDocumentUsingSection
{
    class Program
    {
        static void Main()
        {
            // Load the destination document (the document into which we will insert another document).
            Document destDoc = new Document("Dest.docx");

            // Load the source document (the document whose content will be inserted).
            Document srcDoc = new Document("Source.docx");

            // Create a new empty section in the destination document.
            Section newSection = new Section(destDoc);
            destDoc.Sections.Add(newSection);

            // Insert the content of the first section of the source document
            // into the newly created section of the destination document.
            // Only the body content is copied; page setup, headers, and footers are not.
            newSection.PrependContent(srcDoc.Sections[0]);

            // Save the combined document.
            destDoc.Save("Result.docx");
        }
    }
}
