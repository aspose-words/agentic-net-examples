using System;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Words.Loading;

class ConvertToDoc
{
    static void Main()
    {
        // Path to the source document (any format supported by Aspose.Words).
        string sourcePath = "input.docx";

        // Path where the DOC format document will be saved.
        string targetPath = "output.doc";

        // Load the source document.
        Document doc = new Document(sourcePath);

        // Create save options for the legacy DOC format.
        // The constructor argument ensures the SaveFormat is set to Doc.
        DocSaveOptions saveOptions = new DocSaveOptions(SaveFormat.Doc);

        // Optional: set a password that will be required when loading the saved document.
        saveOptions.Password = "MyPassword";

        // Optional: preserve routing slip information if present.
        saveOptions.SaveRoutingSlip = true;

        // Save the document in DOC format using the specified options.
        doc.Save(targetPath, saveOptions);

        // To demonstrate loading the password‑protected DOC, create LoadOptions with the password.
        LoadOptions loadOptions = new LoadOptions("MyPassword");

        // Load the saved DOC file.
        Document loadedDoc = new Document(targetPath, loadOptions);

        // Output the document text to verify successful conversion.
        Console.WriteLine(loadedDoc.GetText().Trim());
    }
}
