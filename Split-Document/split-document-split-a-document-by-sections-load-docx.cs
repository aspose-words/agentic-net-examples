using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

namespace SplitDocumentBySections
{
    // Implements a callback to customize the filenames of the split parts.
    public class SectionPartRenamer : IDocumentPartSavingCallback
    {
        private readonly string _baseFileName;
        private readonly DocumentSplitCriteria _criteria;
        private int _partIndex = 0;

        public SectionPartRenamer(string baseFileName, DocumentSplitCriteria criteria)
        {
            _baseFileName = baseFileName;
            _criteria = criteria;
        }

        void IDocumentPartSavingCallback.DocumentPartSaving(DocumentPartSavingArgs args)
        {
            // Generate a new filename for each section part.
            string extension = Path.GetExtension(args.DocumentPartFileName);
            string newFileName = $"{Path.GetFileNameWithoutExtension(_baseFileName)}_section_{++_partIndex}{extension}";
            args.DocumentPartFileName = newFileName;

            // Optionally write to a custom stream (here we use the default file handling).
            // If you need a custom stream, uncomment the following lines:
            // string fullPath = Path.Combine(Path.GetDirectoryName(_baseFileName) ?? "", newFileName);
            // args.DocumentPartStream = new FileStream(fullPath, FileMode.Create);
        }
    }

    class Program
    {
        static void Main()
        {
            // Path to the source DOCX file.
            string inputPath = @"C:\Docs\SourceDocument.docx";

            // Base output filename (HTML). Each section will be saved as a separate HTML file.
            string outputPath = @"C:\Docs\SplitDocument.html";

            // Load the document from the DOCX file.
            Document doc = new Document(inputPath);

            // Configure HTML save options to split the document at each section break.
            HtmlSaveOptions saveOptions = new HtmlSaveOptions
            {
                DocumentSplitCriteria = DocumentSplitCriteria.SectionBreak,
                DocumentPartSavingCallback = new SectionPartRenamer(outputPath, DocumentSplitCriteria.SectionBreak)
            };

            // Save the document; Aspose.Words will create separate files for each section.
            doc.Save(outputPath, saveOptions);
        }
    }
}
