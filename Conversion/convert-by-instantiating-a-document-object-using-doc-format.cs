using System.IO;
using Aspose.Words;
using Aspose.Words.Loading;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Path to the source DOC file.
        const string inputPath = @"C:\Docs\source.doc";

        // Open the DOC file as a stream.
        using (FileStream inputStream = new FileStream(inputPath, FileMode.Open, FileAccess.Read))
        {
            // Specify that the document to be loaded is in DOC format.
            LoadOptions loadOptions = new LoadOptions(LoadFormat.Doc, null, null);

            // Instantiate the Document object using the stream and the load options.
            Document doc = new Document(inputStream, loadOptions);

            // (Optional) Save the document back to DOC format using DocSaveOptions.
            const string outputPath = @"C:\Docs\output.doc";
            DocSaveOptions saveOptions = new DocSaveOptions();
            doc.Save(outputPath, saveOptions);
        }
    }
}
