using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

namespace SplitAndMergeExample
{
    // Callback that captures each split part as a separate Document instance.
    public class CollectDocumentPartsCallback : IDocumentPartSavingCallback
    {
        // Stores the cloned part documents for later processing.
        public List<Document> Parts { get; } = new List<Document>();

        // Holds the streams that contain the saved HTML parts.
        private readonly List<MemoryStream> _partStreams = new List<MemoryStream>();

        public void DocumentPartSaving(DocumentPartSavingArgs args)
        {
            // Create a memory stream to capture the HTML part.
            var partStream = new MemoryStream();
            args.DocumentPartStream = partStream;

            // Keep the stream open after the part is saved so we can read it later.
            args.KeepDocumentPartStreamOpen = true;

            // Store the stream for later conversion to a Document.
            _partStreams.Add(partStream);
        }

        // After the main save operation finishes, convert each captured stream into a Document.
        public void LoadParts()
        {
            foreach (var stream in _partStreams)
            {
                stream.Position = 0;
                // Load the HTML part into a Document instance.
                var partDoc = new Document(stream);
                Parts.Add(partDoc);
            }
        }
    }

    public class Program
    {
        public static void Main()
        {
            // -----------------------------------------------------------------
            // 1. Prepare output folder.
            // -----------------------------------------------------------------
            string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
            Directory.CreateDirectory(outputDir);

            // -----------------------------------------------------------------
            // 2. Create a sample document with three sections.
            // -----------------------------------------------------------------
            Document sourceDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(sourceDoc);

            builder.Writeln("Content of Section 1");
            builder.InsertBreak(BreakType.SectionBreakNewPage);
            builder.Writeln("Content of Section 2");
            builder.InsertBreak(BreakType.SectionBreakNewPage);
            builder.Writeln("Content of Section 3");

            // Save the source document (optional, just to see the original).
            sourceDoc.Save(Path.Combine(outputDir, "Source.docx"));

            // -----------------------------------------------------------------
            // 3. Split the document by sections using DocumentSplitCriteria.
            // -----------------------------------------------------------------
            var callback = new CollectDocumentPartsCallback();

            HtmlSaveOptions saveOptions = new HtmlSaveOptions
            {
                DocumentSplitCriteria = DocumentSplitCriteria.SectionBreak,
                DocumentPartSavingCallback = callback
            };

            // The main file name; split parts will be saved alongside it.
            string splitFilePath = Path.Combine(outputDir, "SplitDocument.html");
            sourceDoc.Save(splitFilePath, saveOptions);

            // Convert the captured streams into Document objects.
            callback.LoadParts();

            // Verify that parts were collected.
            if (callback.Parts.Count < 3)
                throw new InvalidOperationException($"Expected 3 parts after splitting, but got {callback.Parts.Count}.");

            // -----------------------------------------------------------------
            // 4. Merge selected parts (first and third sections) into a new document.
            // -----------------------------------------------------------------
            Document mergedDoc = new Document();
            // Remove the default empty section that a new Document creates.
            mergedDoc.RemoveAllChildren();

            // Append the first part (section 1).
            mergedDoc.AppendDocument(callback.Parts[0], ImportFormatMode.KeepSourceFormatting);
            // Append the third part (section 3).
            mergedDoc.AppendDocument(callback.Parts[2], ImportFormatMode.KeepSourceFormatting);

            // Save the merged result.
            string mergedPath = Path.Combine(outputDir, "Merged.docx");
            mergedDoc.Save(mergedPath);

            // -----------------------------------------------------------------
            // 5. Validation – ensure the merged document contains only the expected sections.
            // -----------------------------------------------------------------
            // The merged document should have exactly two sections.
            if (mergedDoc.Sections.Count != 2)
                throw new InvalidOperationException($"Merged document should contain 2 sections, but contains {mergedDoc.Sections.Count}.");

            // Verify content of each section.
            string firstSectionText = mergedDoc.Sections[0].GetText();
            string secondSectionText = mergedDoc.Sections[1].GetText();

            if (!firstSectionText.Contains("Content of Section 1"))
                throw new InvalidOperationException("First merged section does not contain expected text.");

            if (!secondSectionText.Contains("Content of Section 3"))
                throw new InvalidOperationException("Second merged section does not contain expected text.");

            if (firstSectionText.Contains("Content of Section 2") || secondSectionText.Contains("Content of Section 2"))
                throw new InvalidOperationException("Merged document unexpectedly contains content of Section 2.");

            // Program ends successfully.
        }
    }
}
