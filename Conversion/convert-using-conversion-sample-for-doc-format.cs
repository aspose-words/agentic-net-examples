using System;
using Aspose.Words;
using Aspose.Words.Saving;

class ConvertToDocSample
{
    static void Main()
    {
        // Path to the source document (any format supported by Aspose.Words)
        string inputPath = @"C:\Input\SampleDocument.docx";

        // Path where the converted DOC file will be saved
        string outputPath = @"C:\Output\ConvertedDocument.doc";

        // Load the source document using the Document(string) constructor
        Document doc = new Document(inputPath);

        // Option 1: Save directly specifying the SaveFormat enum
        doc.Save(outputPath, SaveFormat.Doc);

        // Option 2: Use a DocSaveOptions object for more control (e.g., password protection)
        // Uncomment the following lines if you need to apply save options.
        /*
        DocSaveOptions saveOptions = new DocSaveOptions(SaveFormat.Doc);
        // Example: protect the saved DOC with a password (optional)
        // saveOptions.Password = "MyPassword";
        doc.Save(outputPath, saveOptions);
        */
    }
}
