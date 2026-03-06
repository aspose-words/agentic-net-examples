using System;
using Aspose.Words;
using Aspose.Words.Saving;

class ProtectOfficeMath
{
    static void Main()
    {
        // Path to the folder that contains the input and output documents.
        string docsPath = @"C:\Docs\";

        // Load the existing DOCX document that contains OfficeMath (equation) objects.
        Document doc = new Document(docsPath + "Input.docx");

        // Protect the entire document as read‑only.
        // This prevents users from editing any content, including OfficeMath objects,
        // while still allowing programmatic modifications via Aspose.Words.
        doc.Protect(ProtectionType.ReadOnly);

        // Save the protected document.
        doc.Save(docsPath + "Input.Protected.docx");
    }
}
