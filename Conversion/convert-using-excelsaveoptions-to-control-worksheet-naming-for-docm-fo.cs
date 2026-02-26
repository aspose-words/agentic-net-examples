using System;
using Aspose.Words;
using Aspose.Words.Saving;

namespace AsposeWordsExample
{
    class Program
    {
        static void Main()
        {
            // Load an existing Word document.
            Document doc = new Document("MyDir/input.docx");

            // Create OoxmlSaveOptions for the DOCM format.
            // The constructor that accepts a SaveFormat ensures the correct options type is used.
            OoxmlSaveOptions saveOptions = new OoxmlSaveOptions(SaveFormat.Docm);

            // Example of setting an additional option – disable embedding the generator name.
            saveOptions.ExportGeneratorName = false;

            // Save the document as a macro‑enabled DOCM file using the specified options.
            doc.Save("ArtifactsDir/output.docm", saveOptions);
        }
    }
}
