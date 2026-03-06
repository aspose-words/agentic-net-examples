using System;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Adjust these paths to point to your actual folders.
        string MyDir = @"C:\Docs\";
        string ArtifactsDir = @"C:\Output\";

        // Load an existing Word document.
        Document doc = new Document(MyDir + "InputDocument.docx");

        // For a macro‑enabled DOCM file we must use OoxmlSaveOptions, not ExcelSaveOptions.
        OoxmlSaveOptions saveOptions = new OoxmlSaveOptions(SaveFormat.Docm);

        // Optional: specify a custom template to be used when the document is saved.
        // saveOptions.DefaultTemplate = MyDir + "CustomTemplate.dotx";

        // Save the document as a DOCM file with the defined options.
        doc.Save(ArtifactsDir + "OutputDocument.docm", saveOptions);
    }
}
