using System;
using Aspose.Words;

class ProtectTxtDocument
{
    static void Main()
    {
        // Load the plain‑text file into an Aspose.Words Document.
        // Aspose.Words automatically creates a single paragraph containing the text.
        Document doc = new Document("input.txt");

        // Protect the entire document but allow editing of form fields only.
        // This sets the protection type to AllowOnlyFormFields.
        doc.Protect(ProtectionType.AllowOnlyFormFields);

        // Save the protected document. The output format can be any supported format,
        // for example DOCX which preserves the protection settings.
        doc.Save("output.docx");
    }
}
