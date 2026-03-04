using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Words.Loading;

class BarcodeDocumentDemo
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Use DocumentBuilder to add content.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Document with a barcode image:");

        // ------------------------------------------------------------
        // Insert a barcode image.
        // For demonstration we use a tiny placeholder PNG encoded in Base64.
        // Replace the Base64 string with an actual barcode image when needed.
        // ------------------------------------------------------------
        const string base64Png = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+XK6cAAAAASUVORK5CYII="; // 1x1 transparent PNG
        byte[] imageBytes = Convert.FromBase64String(base64Png);
        using (MemoryStream imageStream = new MemoryStream(imageBytes))
        {
            builder.InsertImage(imageStream);
        }

        builder.Writeln(); // Add a line break after the image.

        // ------------------------------------------------------------
        // Protect the document (read‑only) with a password.
        // ------------------------------------------------------------
        doc.Protect(ProtectionType.ReadOnly, "protectPwd");

        // ------------------------------------------------------------
        // Encrypt the document on save using a password.
        // ------------------------------------------------------------
        OoxmlSaveOptions saveOptions = new OoxmlSaveOptions
        {
            Password = "encryptPwd"
        };
        string protectedPath = "BarcodeProtectedEncrypted.docx";
        doc.Save(protectedPath, saveOptions);

        // ------------------------------------------------------------
        // Create an edited version of the document for comparison.
        // ------------------------------------------------------------
        Document editedDoc = (Document)doc.Clone(true);
        DocumentBuilder editBuilder = new DocumentBuilder(editedDoc);
        editBuilder.Writeln("Additional line added after the barcode.");

        // ------------------------------------------------------------
        // Load the original (encrypted) document using the password.
        // ------------------------------------------------------------
        LoadOptions loadOptions = new LoadOptions("encryptPwd");
        Document originalDoc = new Document(protectedPath, loadOptions);

        // ------------------------------------------------------------
        // Compare the original and edited documents.
        // ------------------------------------------------------------
        originalDoc.Compare(editedDoc, "Comparer", DateTime.Now);

        // ------------------------------------------------------------
        // Save the comparison result.
        // ------------------------------------------------------------
        string compareResultPath = "ComparisonResult.docx";
        originalDoc.Save(compareResultPath);
    }
}
