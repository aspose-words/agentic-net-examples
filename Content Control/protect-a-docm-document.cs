using System;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Load the existing DOCM document.
        Document doc = new Document("InputDocument.docm");

        // Apply protection. This example uses read‑only protection with a password.
        doc.Protect(ProtectionType.ReadOnly, "MySecretPassword");

        // Save the protected document. The format is inferred from the file extension (.docm).
        doc.Save("ProtectedDocument.docm");
    }
}
