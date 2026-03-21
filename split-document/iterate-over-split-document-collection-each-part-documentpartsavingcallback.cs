using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Resolve paths relative to the executable directory.
        string baseDir = AppContext.BaseDirectory;
        string dataDir = Path.Combine(baseDir, "Data");
        string inputPath = Path.Combine(dataDir, "Input.docx");
        string outputFolder = Path.Combine(baseDir, "Output");

        // Ensure required directories exist.
        Directory.CreateDirectory(dataDir);
        Directory.CreateDirectory(outputFolder);

        // If the input document does not exist, create a simple one.
        if (!File.Exists(inputPath))
        {
            var tempDoc = new Document();
            var builder = new DocumentBuilder(tempDoc);
            builder.Writeln("This is a sample document.");
            builder.InsertBreak(BreakType.SectionBreakNewPage);
            builder.Writeln("Second section.");
            tempDoc.Save(inputPath);
        }

        // Load the source document.
        Document doc = new Document(inputPath);

        // Base file name for the main HTML file.
        string baseFileName = "SplitDocument.html";

        // Configure HTML save options to split the document by section.
        HtmlSaveOptions saveOptions = new HtmlSaveOptions
        {
            DocumentSplitCriteria = DocumentSplitCriteria.SectionBreak,
            DocumentPartSavingCallback = new SavedDocumentPartRename(baseFileName, outputFolder, DocumentSplitCriteria.SectionBreak)
        };

        // Save the document; the callback will be invoked for each part.
        doc.Save(Path.Combine(outputFolder, baseFileName), saveOptions);
    }
}

// Callback that controls how each document part is saved.
class SavedDocumentPartRename : IDocumentPartSavingCallback
{
    private readonly string _baseFileName;
    private readonly string _outputFolder;
    private readonly DocumentSplitCriteria _criteria;
    private int _partIndex;

    public SavedDocumentPartRename(string baseFileName, string outputFolder, DocumentSplitCriteria criteria)
    {
        _baseFileName = baseFileName;
        _outputFolder = outputFolder;
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

        // Build a unique file name for the part.
        string partFileName = $"{Path.GetFileNameWithoutExtension(_baseFileName)}_Part{++_partIndex}_{partType}{Path.GetExtension(args.DocumentPartFileName)}";

        // Set the file name (Aspose will use this name when saving the part).
        args.DocumentPartFileName = partFileName;

        // Provide a stream so the part is written directly to the desired location.
        args.DocumentPartStream = new FileStream(Path.Combine(_outputFolder, partFileName), FileMode.Create);
        args.KeepDocumentPartStreamOpen = false;
    }
}
