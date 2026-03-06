using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

class ProtectDocxExample
{
    static void Main()
    {
        // Define the folder where the document will be saved.
        string artifactsDir = Path.Combine(Environment.CurrentDirectory, "Output");
        Directory.CreateDirectory(artifactsDir);

        // Create a new blank document.
        Document doc = new Document();

        // Add some content to the document.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("This is a protected DOCX document.");

        // Apply read‑only protection with a password.
        doc.Protect(ProtectionType.ReadOnly, "MyPassword");

        // Optionally, encrypt the file with a password using OOXML encryption.
        OoxmlSaveOptions saveOptions = new OoxmlSaveOptions
        {
            Password = "MyPassword"
        };

        // Save the protected document as DOCX.
        string outputPath = Path.Combine(artifactsDir, "ProtectedDocument.docx");
        doc.Save(outputPath, saveOptions);

        // Inform the user.
        Console.WriteLine($"Document saved and protected at: {outputPath}");
    }
}
