using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.DigitalSignatures;
using Aspose.Words.Properties;

class DocmPrerequisitesDemo
{
    static void Main()
    {
        // Path to the source DOCM file.
        string srcPath = @"C:\Docs\Sample.docm";

        // Detect the file format and its characteristics without loading the whole document.
        FileFormatInfo formatInfo = FileFormatUtil.DetectFileFormat(srcPath);

        // Output basic format information.
        Console.WriteLine($"Load format: {formatInfo.LoadFormat}");
        Console.WriteLine($"Is encrypted: {formatInfo.IsEncrypted}");
        Console.WriteLine($"Has macros: {formatInfo.HasMacros}");
        Console.WriteLine($"Has digital signature: {formatInfo.HasDigitalSignature}");

        // Load the document only if it is not encrypted (otherwise a password is required).
        Document doc = null;
        if (!formatInfo.IsEncrypted)
        {
            doc = new Document(srcPath);
        }
        else
        {
            // Example of loading an encrypted document (password would be supplied here).
            // LoadOptions loadOptions = new LoadOptions("password");
            // doc = new Document(srcPath, loadOptions);
            Console.WriteLine("Document is encrypted; loading with password is not demonstrated.");
            return;
        }

        // Verify macro presence via the Document property (more reliable after loading).
        Console.WriteLine($"Document.HasMacros (after load): {doc.HasMacros}");

        // Check for digital signatures using the Document's collection.
        DigitalSignatureCollection signatures = doc.DigitalSignatures;
        Console.WriteLine($"Digital signatures count: {signatures.Count}");
        Console.WriteLine($"All signatures valid: {signatures.IsValid}");

        // Example: if the document is not signed, sign it with a certificate.
        if (signatures.Count == 0)
        {
            // Path to a PKCS#12 certificate file that contains a private key.
            string certPath = @"C:\Certs\mycert.pfx";
            string certPassword = "certPassword";

            // Create a CertificateHolder from the certificate file.
            CertificateHolder certHolder = CertificateHolder.Create(certPath, certPassword);

            // Define signing options (optional).
            SignOptions signOptions = new SignOptions
            {
                Comments = "Signed by DocmPrerequisitesDemo",
                SignTime = DateTime.Now
            };

            // Sign the document and overwrite the original file.
            DigitalSignatureUtil.Sign(srcPath, srcPath, certHolder, signOptions);

            // Re‑detect format info to confirm the signature was added.
            FileFormatInfo postSignInfo = FileFormatUtil.DetectFileFormat(srcPath);
            Console.WriteLine($"Has digital signature after signing: {postSignInfo.HasDigitalSignature}");
        }

        // Example of checking a built‑in property (SharedDocument) which may be relevant for DOCM files.
        bool isShared = doc.BuiltInDocumentProperties.SharedDocument;
        Console.WriteLine($"Built‑in property SharedDocument: {isShared}");

        // Save a copy of the document (demonstrating the allowed save rule).
        string copyPath = @"C:\Docs\Sample_Copy.docm";
        doc.Save(copyPath);
        Console.WriteLine($"Document saved to: {copyPath}");
    }
}
