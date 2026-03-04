using Aspose.Words;
using Aspose.Words.Settings;

class Program
{
    static void Main()
    {
        // Load an existing DOCM file or create a new document.
        // Replace "input.docm" with the path to your source DOCM if needed.
        Document doc = new Document(); // new Document("input.docm");

        // Apply read‑only protection with a password.
        doc.Protect(ProtectionType.ReadOnly, "MyPassword");

        // Optionally, also set write‑protection settings.
        doc.WriteProtection.SetPassword("MyPassword");
        doc.WriteProtection.ReadOnlyRecommended = true;

        // Save the protected document as DOCM.
        doc.Save("ProtectedDocument.docm");
    }
}
