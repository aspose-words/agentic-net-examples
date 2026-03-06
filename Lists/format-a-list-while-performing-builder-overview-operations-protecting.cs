using System;
using Aspose.Words;
using Aspose.Words.Lists;
using Aspose.Words.Saving;

class DocumentProcessing
{
    static void Main()
    {
        // Load the source DOCX document.
        Document originalDoc = new Document("Original.docx");

        // ---------- Builder Overview: format a numbered list ----------
        DocumentBuilder builder = new DocumentBuilder(originalDoc);
        // Ensure we start on a new paragraph.
        builder.Writeln();

        // Create a new list based on the default numbered template.
        builder.ListFormat.List = originalDoc.Lists.Add(ListTemplate.NumberDefault);

        // Add several list items.
        for (int i = 1; i <= 5; i++)
        {
            builder.Writeln($"Item {i}");
        }

        // End the list formatting.
        builder.ListFormat.RemoveNumbers();

        // ---------- Protect (read‑only) the document ----------
        // Apply read‑only protection with a password.
        originalDoc.Protect(ProtectionType.ReadOnly, "protectPwd");

        // ---------- Encrypt the document on save ----------
        // Use OoxmlSaveOptions to set an encryption password.
        OoxmlSaveOptions saveOptions = new OoxmlSaveOptions(SaveFormat.Docx)
        {
            Password = "encryptPwd"
        };
        originalDoc.Save("Original_Protected_Encrypted.docx", saveOptions);

        // ---------- Compare with another document ----------
        // Load the document to compare against.
        Document editedDoc = new Document("Edited.docx");

        // Ensure both documents have no revisions before comparison.
        if (originalDoc.Revisions.Count == 0 && editedDoc.Revisions.Count == 0)
        {
            originalDoc.Compare(editedDoc, "Comparer", DateTime.Now);
        }

        // Optionally accept all revisions so the original becomes identical to the edited version.
        originalDoc.Revisions.AcceptAll();

        // Save the comparison result.
        originalDoc.Save("ComparisonResult.docx");
    }
}
