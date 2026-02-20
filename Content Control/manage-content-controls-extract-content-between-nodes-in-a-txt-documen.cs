using System;
using Aspose.Words;
using Aspose.Words.Markup;

class Program
{
    static void Main()
    {
        // Load the plain‑text document. Aspose.Words automatically detects the TXT format.
        Document doc = new Document("input.txt");

        // Try to locate a content control (structured document tag) by title.
        // GetByTitle returns IStructuredDocumentTag, so we need an explicit cast to StructuredDocumentTag.
        StructuredDocumentTag sdt = doc.Range.StructuredDocumentTags.GetByTitle("MyTagTitle") as StructuredDocumentTag;

        // If the tag was not found, fall back to the first tag in the collection (if any).
        if (sdt == null && doc.Range.StructuredDocumentTags.Count > 0)
        {
            sdt = doc.Range.StructuredDocumentTags[0] as StructuredDocumentTag;
        }

        // Extract the text that lies inside the content control.
        string extractedText = sdt?.Range?.Text ?? string.Empty;

        // Output the extracted text.
        Console.WriteLine(extractedText);
    }
}
