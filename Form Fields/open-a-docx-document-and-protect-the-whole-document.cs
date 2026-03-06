using System;
using Aspose.Words;

class Program
{
    static void Main()
    {
        // Load the existing DOCX file.
        Document doc = new Document("input.docx");

        // Protect the entire document as read‑only.
        // This prevents any changes unless the protection is removed programmatically.
        doc.Protect(ProtectionType.ReadOnly);

        // Save the protected document.
        doc.Save("protected.docx");
    }
}
