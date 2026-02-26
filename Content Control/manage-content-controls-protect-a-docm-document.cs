using System;
using Aspose.Words;

class ProtectDocmExample
{
    static void Main()
    {
        // Path to the source DOCM file (must exist on disk).
        string inputPath = @"C:\Docs\Sample.docm";

        // Path where the protected DOCM will be saved.
        string outputPath = @"C:\Docs\Sample.Protected.docm";

        // Load the existing DOCM document.
        Document doc = new Document(inputPath);

        // Apply protection that allows only form field editing.
        // A password is optional; an empty string means no password.
        doc.Protect(ProtectionType.AllowOnlyFormFields, "MySecretPassword");

        // Save the protected document back to DOCM format.
        doc.Save(outputPath);
    }
}
