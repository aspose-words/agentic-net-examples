using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    // Entry point of the console application.
    public static void Main()
    {
        // Define a folder that represents the network share where split parts will be saved.
        // In a real scenario this could be a UNC path like @"\\ServerName\ShareFolder".
        string networkSharePath = Path.Combine(Environment.CurrentDirectory, "NetworkShare");
        Directory.CreateDirectory(networkSharePath);

        // Create a sample document with three sections.
        Document doc = CreateSampleDocument();

        // Configure HTML save options to split the document by section breaks.
        HtmlSaveOptions saveOptions = new HtmlSaveOptions
        {
            DocumentSplitCriteria = DocumentSplitCriteria.SectionBreak,
            // The base file name is required; the callback will override the actual part names.
            // It can be any valid file name; the parts will be saved via the callback.
            DocumentPartSavingCallback = new SectionSplitCallback(networkSharePath, "SplitDocument")
        };

        // Save the document. The callback will write each section to a separate file in the network share.
        string dummyOutputPath = Path.Combine(networkSharePath, "SplitDocument.html");
        doc.Save(dummyOutputPath, saveOptions);

        // Verify that the expected number of split files were created.
        int expectedParts = doc.Sections.Count;
        int actualParts = Directory.GetFiles(networkSharePath, "SplitDocument_part*.html").Length;

        if (actualParts != expectedParts)
        {
            throw new InvalidOperationException(
                $"Expected {expectedParts} split parts, but found {actualParts} in the network share.");
        }

        // Optionally, indicate success (no interactive output required).
    }

    // Creates a document containing three sections with simple text.
    private static Document CreateSampleDocument()
    {
        Document document = new Document();
        DocumentBuilder builder = new DocumentBuilder(document);

        builder.Writeln("Section 1 - First paragraph.");
        builder.InsertBreak(BreakType.SectionBreakNewPage);

        builder.Writeln("Section 2 - First paragraph.");
        builder.InsertBreak(BreakType.SectionBreakNewPage);

        builder.Writeln("Section 3 - First paragraph.");

        return document;
    }

    // Callback that redirects each split part to a file in the specified network share.
    private class SectionSplitCallback : IDocumentPartSavingCallback
    {
        private readonly string _sharePath;
        private readonly string _baseFileName;
        private int _partIndex = 0;

        public SectionSplitCallback(string sharePath, string baseFileName)
        {
            _sharePath = sharePath;
            _baseFileName = baseFileName;
        }

        void IDocumentPartSavingCallback.DocumentPartSaving(DocumentPartSavingArgs args)
        {
            // Increment part counter.
            _partIndex++;

            // Build a unique file name for this part.
            string partFileName = $"{_baseFileName}_part{_partIndex}.html";

            // Set the stream to write the part directly to the network share.
            string fullPath = Path.Combine(_sharePath, partFileName);
            args.DocumentPartStream = new FileStream(fullPath, FileMode.Create);
            // Ensure Aspose.Words closes the stream after writing.
            args.KeepDocumentPartStreamOpen = false;
        }
    }
}
