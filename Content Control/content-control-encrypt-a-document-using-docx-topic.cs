using System;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Words.Loading;
using Aspose.Words;

namespace AsposeWordsEncryptionDemo
{
    class Program
    {
        static void Main()
        {
            // Path where the encrypted document will be saved.
            string outputPath = "EncryptedDocument.docx";

            // 1. Create a new blank document.
            Document doc = new Document();

            // 2. Add some content using DocumentBuilder.
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.Writeln("Hello world! This document is encrypted.");

            // 3. Configure OoxmlSaveOptions to encrypt the document with a password.
            OoxmlSaveOptions saveOptions = new OoxmlSaveOptions();
            saveOptions.Password = "MySecretPassword";

            // 4. Save the document using the save options (encryption applied).
            doc.Save(outputPath, saveOptions);

            // -----------------------------------------------------------------
            // Demonstrate that opening the file without a password throws an exception.
            try
            {
                // This will throw IncorrectPasswordException because no password is supplied.
                Document loadFail = new Document(outputPath);
                Console.WriteLine("Unexpected: Document opened without password.");
            }
            catch (IncorrectPasswordException)
            {
                Console.WriteLine("Correctly caught IncorrectPasswordException when opening without password.");
            }

            // -----------------------------------------------------------------
            // Load the encrypted document by providing the correct password via LoadOptions.
            LoadOptions loadOptions = new LoadOptions("MySecretPassword");
            Document loadedDoc = new Document(outputPath, loadOptions);

            // Verify that the text was loaded correctly.
            string text = loadedDoc.GetText().Trim();
            Console.WriteLine($"Loaded document text: \"{text}\"");
        }
    }
}
