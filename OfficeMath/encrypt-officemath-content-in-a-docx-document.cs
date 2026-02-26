using System;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Words.Loading;

class EncryptOfficeMathExample
{
    static void Main()
    {
        // Path where the document will be saved.
        string artifactsDir = @"C:\Temp\Artifacts\";

        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add some text and an equation (OfficeMath) to the document.
        // In a real scenario you would insert an actual OfficeMath object.
        builder.Writeln("This paragraph contains an equation:");
        builder.Writeln("x² + y² = z²"); // Placeholder for OfficeMath content.

        // Configure save options to encrypt the document with a password.
        OoxmlSaveOptions saveOptions = new OoxmlSaveOptions();
        saveOptions.Password = "MySecretPassword";

        // Save the encrypted DOCX file.
        string encryptedPath = System.IO.Path.Combine(artifactsDir, "EncryptedOfficeMath.docx");
        doc.Save(encryptedPath, saveOptions);

        // Load the encrypted document using the same password.
        LoadOptions loadOptions = new LoadOptions("MySecretPassword");
        Document loadedDoc = new Document(encryptedPath, loadOptions);

        // Output the document text to verify it was loaded correctly.
        Console.WriteLine(loadedDoc.GetText().Trim());
    }
}
