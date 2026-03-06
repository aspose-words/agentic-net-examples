using System;
using Aspose.Words;
using Aspose.Words.Loading;
using Aspose.Words.Saving;

class EncryptOfficeMathExample
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add some text and an OfficeMath equation.
        builder.Writeln("Below is an equation:");
        // Insert a simple equation (a² + b² = c²) using a field that Word renders as an equation.
        // Aspose.Words does not expose a direct InsertEquation method; the equivalent is to insert an
        // EQ field which Word displays as an OfficeMath object.
        builder.InsertField("EQ \\o(a,2)+\\o(b,2)=\\o(c,2)");
        builder.Writeln();

        // Configure save options to encrypt the document with a password using the ECMA376 standard encryption algorithm.
        OoxmlSaveOptions saveOptions = new OoxmlSaveOptions();
        saveOptions.Password = "Secret123";

        // Save the encrypted document.
        string filePath = "EncryptedOfficeMath.docx";
        doc.Save(filePath, saveOptions);

        // Attempt to load the encrypted document without a password.
        // This should throw an IncorrectPasswordException.
        try
        {
            Document _ = new Document(filePath);
        }
        catch (IncorrectPasswordException)
        {
            // Expected – the document is password protected.
            Console.WriteLine("Document is password protected – cannot open without password.");
        }

        // Load the encrypted document using the correct password.
        LoadOptions loadOptions = new LoadOptions("Secret123");
        Document loadedDoc = new Document(filePath, loadOptions);

        // Verify that the document content (including the equation) is accessible.
        Console.WriteLine(loadedDoc.GetText());
    }
}
