using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a sample document with three pages.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Page 1 content.");
        builder.InsertBreak(BreakType.PageBreak);
        builder.Writeln("Page 2 content.");
        builder.InsertBreak(BreakType.PageBreak);
        builder.Writeln("Page 3 content.");

        // Set up HTML save options to split the document at page breaks.
        string outputBaseName = "SplitDocument.html";
        HtmlSaveOptions saveOptions = new HtmlSaveOptions
        {
            DocumentSplitCriteria = DocumentSplitCriteria.PageBreak,
            DocumentPartSavingCallback = new SavedDocumentPartRename(outputBaseName)
        };

        // Save the document; the callback will create separate HTML files for each page.
        doc.Save(outputBaseName, saveOptions);

        // Verify that the expected number of page files were created.
        string outputDirectory = Path.GetDirectoryName(Path.GetFullPath(outputBaseName));
        string[] pageFiles = Directory.GetFiles(outputDirectory, "SplitDocument.Page_*.html")
                                      .OrderBy(f => f)
                                      .ToArray();

        if (pageFiles.Length != doc.PageCount)
            throw new InvalidOperationException($"Expected {doc.PageCount} page files, but found {pageFiles.Length}.");

        // Output the list of generated files.
        foreach (string file in pageFiles)
            Console.WriteLine($"Generated: {Path.GetFileName(file)}");
    }

    // Callback that assigns a unique file name for each document part (page).
    private class SavedDocumentPartRename : IDocumentPartSavingCallback
    {
        private int _count = 0;
        private readonly string _baseFileName;

        public SavedDocumentPartRename(string baseFileName)
        {
            _baseFileName = baseFileName;
        }

        void IDocumentPartSavingCallback.DocumentPartSaving(DocumentPartSavingArgs args)
        {
            _count++;
            // Create a deterministic file name for each page.
            string outFileName = $"{Path.GetFileNameWithoutExtension(_baseFileName)}.Page_{_count}.html";

            // Set the file name and stream for the page.
            args.DocumentPartFileName = outFileName;
            args.DocumentPartStream = new FileStream(outFileName, FileMode.Create);
            args.KeepDocumentPartStreamOpen = false;
        }
    }
}
