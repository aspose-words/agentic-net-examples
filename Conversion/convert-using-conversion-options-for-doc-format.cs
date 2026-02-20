using System;
using Aspose.Words;
using Aspose.Words.Loading;
using Aspose.Words.Saving;

class DocConversionExample
{
    static void Main()
    {
        // Path to the source document (any format supported by Aspose.Words)
        string inputPath = @"C:\Docs\source.docx";

        // Path where the DOC format document will be saved
        string outputPath = @"C:\Docs\converted.doc";

        // ---------- Load the document ----------
        // Create LoadOptions to customize loading (e.g., password, base URI, etc.)
        LoadOptions loadOptions = new LoadOptions();
        // Example: set a password if the source document is encrypted
        // loadOptions.Password = "sourcePassword";

        // Load the document using the specified options
        Document doc = new Document(inputPath, loadOptions);

        // ---------- Configure DOC save options ----------
        DocSaveOptions saveOptions = new DocSaveOptions();
        // Ensure fields are updated before saving
        saveOptions.UpdateFields = true;
        // Embed the Aspose.Words generator name/version into the output file
        saveOptions.ExportGeneratorName = true;
        // Specify the exact format to save as DOC (optional, as DocSaveOptions defaults to DOC)
        saveOptions.SaveFormat = SaveFormat.Doc;
        // Example: set a password to encrypt the output DOC file
        // saveOptions.Password = "outputPassword";

        // ---------- Save the document as DOC ----------
        doc.Save(outputPath, saveOptions);
    }
}
