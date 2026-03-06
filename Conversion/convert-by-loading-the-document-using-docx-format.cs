using System;
using Aspose.Words;
using Aspose.Words.Loading;

class Program
{
    static void Main()
    {
        // Path to the source DOCX file.
        string sourcePath = "input.docx";

        // Convert the SaveFormat.Docx enum to the corresponding LoadFormat.
        LoadFormat docxLoadFormat = FileFormatUtil.SaveFormatToLoadFormat(SaveFormat.Docx);

        // Create LoadOptions and explicitly set the format to DOCX.
        LoadOptions loadOptions = new LoadOptions
        {
            LoadFormat = docxLoadFormat
        };

        // Load the document using the specified DOCX load format.
        Document doc = new Document(sourcePath, loadOptions);

        // (Optional) Save the document to verify it was loaded correctly.
        string outputPath = "output.docx";
        doc.Save(outputPath);
    }
}
