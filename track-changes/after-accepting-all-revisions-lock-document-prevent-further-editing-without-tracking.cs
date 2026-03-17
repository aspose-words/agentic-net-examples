using System;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Load an existing document (replace with your actual file path)
        Document doc = new Document("input.docx");

        // Accept all tracked changes in the document
        doc.AcceptAllRevisions();

        // Protect the document so that any further editing is only allowed as revisions
        // This prevents normal editing without tracking changes.
        doc.Protect(ProtectionType.AllowOnlyRevisions);

        // Save the locked document (replace with your desired output path)
        doc.Save("output.docx");
    }
}
