using System;
using Aspose.Words;
using Aspose.Words.Saving;

class ProtectOfficeMath
{
    static void Main()
    {
        // Load the existing DOCX document that contains OfficeMath objects.
        Document doc = new Document("InputWithOfficeMath.docx");

        // Apply read‑only protection to the whole document.
        // This prevents any changes (including edits to OfficeMath) when the file is opened in Microsoft Word.
        doc.Protect(ProtectionType.ReadOnly);

        // Save the protected document.
        doc.Save("OutputProtected.docx");
    }
}
