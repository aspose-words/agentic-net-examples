using System;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Words.Vba;
using Aspose.Words.Loading;

class Program
{
    static void Main()
    {
        // Paths to the source documents.
        string sourcePath = "DocumentWithMacros.docm";   // Document that contains VBA macros.
        string comparePath = "AnotherDocument.docx";    // Document to compare against.
        string protectedPath = "ProtectedDocument.docx"; // Output path for the protected/encrypted file.

        // -----------------------------------------------------------------
        // 1. Load the source document (it may contain macros).
        // -----------------------------------------------------------------
        Document sourceDoc = new Document(sourcePath);

        // -----------------------------------------------------------------
        // 2. Apply write‑protection (password required to modify) and
        //    encrypt the file with a password when saving.
        // -----------------------------------------------------------------
        // Write‑protection – does NOT encrypt the content.
        sourceDoc.WriteProtection.SetPassword("WritePass");
        sourceDoc.WriteProtection.ReadOnlyRecommended = true;

        // Encryption – set a password in OoxmlSaveOptions.
        OoxmlSaveOptions saveOptions = new OoxmlSaveOptions(SaveFormat.Docx)
        {
            Password = "EncryptPass"
        };

        // Save the protected and encrypted document.
        sourceDoc.Save(protectedPath, saveOptions);

        // -----------------------------------------------------------------
        // 3. Load the protected document back (providing the encryption password).
        // -----------------------------------------------------------------
        LoadOptions loadOptions = new LoadOptions("EncryptPass");
        Document protectedDoc = new Document(protectedPath, loadOptions);

        // -----------------------------------------------------------------
        // 4. Load the document to compare with.
        // -----------------------------------------------------------------
        Document compareDoc = new Document(comparePath);

        // -----------------------------------------------------------------
        // 5. Compare the two documents.
        // -----------------------------------------------------------------
        bool textsEqual = protectedDoc.GetText() == compareDoc.GetText();

        bool macrosEqual = protectedDoc.HasMacros == compareDoc.HasMacros &&
                           (!protectedDoc.HasMacros || 
                            protectedDoc.VbaProject.Modules.Count == compareDoc.VbaProject.Modules.Count);

        bool writeProtectionEqual = protectedDoc.WriteProtection.IsWriteProtected == compareDoc.WriteProtection.IsWriteProtected &&
                                    (!protectedDoc.WriteProtection.IsWriteProtected ||
                                     protectedDoc.WriteProtection.ValidatePassword("WritePass") ==
                                     compareDoc.WriteProtection.ValidatePassword("WritePass"));

        // -----------------------------------------------------------------
        // 6. Detect encryption status using FileFormatUtil.
        // -----------------------------------------------------------------
        FileFormatInfo formatInfo = FileFormatUtil.DetectFileFormat(protectedPath);
        bool isEncrypted = formatInfo.IsEncrypted;

        // -----------------------------------------------------------------
        // 7. Output the comparison results.
        // -----------------------------------------------------------------
        Console.WriteLine($"Texts equal: {textsEqual}");
        Console.WriteLine($"Macros equal: {macrosEqual}");
        Console.WriteLine($"Write protection equal: {writeProtectionEqual}");
        Console.WriteLine($"File is encrypted: {isEncrypted}");
    }
}
