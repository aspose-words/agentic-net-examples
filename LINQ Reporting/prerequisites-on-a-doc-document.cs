using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Words.Loading;
using Aspose.Words.Settings;
using Aspose.Words.DigitalSignatures;
using Aspose.Words.Properties;

class Program
{
    static void Main()
    {
        // 1. Create a blank DOC document.
        Document doc = new Document();

        // 2. Add minimal content using DocumentBuilder.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Prerequisites demonstration.");

        // 3. Optimize the document for a specific Word version (e.g., Word 2007).
        doc.CompatibilityOptions.OptimizeFor(MsWordVersion.Word2007);

        // 4. Save the document as a legacy .doc file with a password.
        DocSaveOptions saveOptions = new DocSaveOptions(SaveFormat.Doc);
        saveOptions.Password = "Secret123";
        string docPath = "Prerequisites.doc";
        doc.Save(docPath, saveOptions);

        // 5. Load the password‑protected document using LoadOptions.
        LoadOptions loadOptions = new LoadOptions("Secret123");
        Document loadedDoc = new Document(docPath, loadOptions);

        // 6. Detect the file format and encryption status of the saved file.
        FileFormatInfo formatInfo = FileFormatUtil.DetectFileFormat(docPath);
        Console.WriteLine($"LoadFormat: {formatInfo.LoadFormat}, IsEncrypted: {formatInfo.IsEncrypted}");

        // 7. Access a built‑in document property (read‑only example).
        bool isShared = loadedDoc.BuiltInDocumentProperties.SharedDocument;
        Console.WriteLine($"SharedDocument property: {isShared}");

        // 8. (Optional) Digitally sign the document if a certificate is available.
        string certPath = "mycert.pfx";
        if (File.Exists(certPath))
        {
            // Load the certificate (replace "certPassword" with the actual password).
            CertificateHolder certHolder = CertificateHolder.Create(certPath, "certPassword");
            SignOptions signOptions = new SignOptions { SignTime = DateTime.Now };
            string signedPath = "Prerequisites_Signed.doc";

            // Sign the original .doc file and write the signed version.
            DigitalSignatureUtil.Sign(docPath, signedPath, certHolder, signOptions);
            Console.WriteLine("Document signed successfully.");
        }
    }
}
