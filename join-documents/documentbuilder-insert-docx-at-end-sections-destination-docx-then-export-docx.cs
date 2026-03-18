using System;
using Aspose.Words;

class InsertDocAtSectionEnds
{
    static void Main()
    {
        // Load the destination document.
        Document dstDoc = new Document("Destination.docx");

        // Load the source document that will be inserted.
        Document srcDoc = new Document("Source.docx");

        // For each section in the destination document, append the content of every
        // section from the source document to the end of that destination section.
        foreach (Section destSection in dstDoc.Sections)
        {
            foreach (Section srcSection in srcDoc.Sections)
            {
                // AppendContent copies only the body content of the source section.
                destSection.AppendContent(srcSection);
            }
        }

        // Save the modified document.
        dstDoc.Save("Result.docx");
    }
}
