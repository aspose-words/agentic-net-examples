using System;
using Aspose.Words;
using Aspose.Words.Loading;

class Program
{
    static void Main()
    {
        // Path to the source document (DOC format)
        string sourcePath = "source.doc";

        // Specify load options and explicitly set the format to DOC
        LoadOptions loadOptions = new LoadOptions
        {
            LoadFormat = LoadFormat.Doc
        };

        // Load the document using the specified options
        Document doc = new Document(sourcePath, loadOptions);

        // Example: save the loaded document to another format (e.g., DOCX)
        string outputPath = "output.docx";
        doc.Save(outputPath);
    }
}
