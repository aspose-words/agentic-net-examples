using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class DotmSplitter
{
    /// <summary>
    /// Splits a DOTM (macro‑enabled template) into separate HTML parts.
    /// Each part is saved as a distinct file using a custom callback.
    /// </summary>
    /// <param name="inputPath">Full path to the source .dotm file.</param>
    /// <param name="outputFolder">Folder where the split parts will be written.</param>
    /// <param name="baseFileName">Base name used for the generated files (without extension).</param>
    public static void SplitDotm(string inputPath, string outputFolder, string baseFileName)
    {
        // Ensure the output directory exists.
        Directory.CreateDirectory(outputFolder);

        // Load the DOTM document.
        Document doc = new Document(inputPath);

        // Configure HTML save options to split the document at each section break.
        HtmlSaveOptions options = new HtmlSaveOptions
        {
            DocumentSplitCriteria = DocumentSplitCriteria.SectionBreak,
            // Assign a callback that defines how each part is named and where it is saved.
            DocumentPartSavingCallback = new SavedDocumentPartRename(outputFolder, baseFileName, DocumentSplitCriteria.SectionBreak)
        };

        // Save the main HTML file; the callback will handle the individual parts.
        string mainFilePath = Path.Combine(outputFolder, baseFileName + ".html");
        doc.Save(mainFilePath, options);
    }

    /// <summary>
    /// Callback that customizes the filename and stream for each document part.
    /// Implements IDocumentPartSavingCallback as required by Aspose.Words.
    /// </summary>
    private class SavedDocumentPartRename : IDocumentPartSavingCallback
    {
        private readonly string _outputFolder;
        private readonly string _baseFileName;
        private readonly DocumentSplitCriteria _criteria;
        private int _count;

        public SavedDocumentPartRename(string outputFolder, string baseFileName, DocumentSplitCriteria criteria)
        {
            _outputFolder = outputFolder;
            _baseFileName = baseFileName;
            _criteria = criteria;
        }

        void IDocumentPartSavingCallback.DocumentPartSaving(DocumentPartSavingArgs args)
        {
            // Determine a readable part type based on the split criteria.
            string partType = _criteria switch
            {
                DocumentSplitCriteria.PageBreak => "Page",
                DocumentSplitCriteria.ColumnBreak => "Column",
                DocumentSplitCriteria.SectionBreak => "Section",
                DocumentSplitCriteria.HeadingParagraph => "Heading",
                _ => "Part"
            };

            // Build a unique filename for this part.
            string partFileName = $"{_baseFileName}_part{++_count}_{partType}{Path.GetExtension(args.DocumentPartFileName)}";

            // Set the filename (without path). Aspose will use the folder of the main file.
            args.DocumentPartFileName = partFileName;

            // Alternatively, write directly to a stream in the output folder.
            string fullPath = Path.Combine(_outputFolder, partFileName);
            args.DocumentPartStream = new FileStream(fullPath, FileMode.Create);

            // Keep the stream closed after Aspose writes (default behavior).
            args.KeepDocumentPartStreamOpen = false;
        }
    }
}

// Example usage.
class Program
{
    static void Main()
    {
        string inputDotm = @"C:\Docs\Template.dotm";
        string outputFolder = @"C:\Docs\SplitOutput";
        string baseFileName = "TemplateSplit";

        DotmSplitter.SplitDotm(inputDotm, outputFolder, baseFileName);
    }
}
