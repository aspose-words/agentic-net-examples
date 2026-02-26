using System;
using Aspose.Words;
using Aspose.Words.Saving;

class RemoveHeadersFootersAndSaveAsTemplate
{
    static void Main()
    {
        // Path to the source DOC file.
        string sourcePath = @"MyDir\DocumentWithHeadersFooters.doc";

        // Load the existing document.
        Document doc = new Document(sourcePath);

        // Remove all headers and footers from every section.
        foreach (Section section in doc.Sections)
        {
            // This clears the content of headers/footers but keeps the objects,
            // effectively unlinking them from the previous sections.
            section.ClearHeadersFooters();
        }

        // Prepare save options for the DOT (Word template) format.
        DocSaveOptions saveOptions = new DocSaveOptions(SaveFormat.Dot);

        // Save the modified document as a DOT file.
        string outputPath = @"ArtifactsDir\DocumentWithoutHeadersFooters.dot";
        doc.Save(outputPath, saveOptions);
    }
}
