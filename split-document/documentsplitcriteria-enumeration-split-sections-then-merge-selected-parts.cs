using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // 1. Create a source document with three sections.
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);
        builder.Writeln("Content of Section 1");
        builder.InsertBreak(BreakType.SectionBreakNewPage);
        builder.Writeln("Content of Section 2");
        builder.InsertBreak(BreakType.SectionBreakNewPage);
        builder.Writeln("Content of Section 3");

        // 2. Save the document as HTML, splitting it at each section.
        string baseFileName = "SplitDocument.html";
        HtmlSaveOptions saveOptions = new HtmlSaveOptions();
        saveOptions.DocumentSplitCriteria = DocumentSplitCriteria.SectionBreak;
        saveOptions.DocumentPartSavingCallback = new SavedDocumentPartRename(baseFileName, DocumentSplitCriteria.SectionBreak);
        sourceDoc.Save(baseFileName, saveOptions);

        // 3. Load the desired parts (for example, part 1 and part 3).
        Document part1 = new Document("SplitDocument part 1, of type Section.html");
        Document part3 = new Document("SplitDocument part 3, of type Section.html");

        // 4. Merge the selected parts into a new document.
        Document mergedDoc = new Document();
        // Ensure the merged document has at least one empty section.
        mergedDoc.RemoveAllChildren();
        mergedDoc.AppendChild(new Section(mergedDoc));

        mergedDoc.AppendDocument(part1, ImportFormatMode.KeepSourceFormatting);
        mergedDoc.AppendDocument(part3, ImportFormatMode.KeepSourceFormatting);

        // 5. Save the merged result.
        mergedDoc.Save("MergedSelectedParts.docx");
    }

    // Callback that assigns custom filenames to each split part.
    private class SavedDocumentPartRename : IDocumentPartSavingCallback
    {
        private readonly string _baseFileName;
        private readonly DocumentSplitCriteria _criteria;
        private int _partIndex;

        public SavedDocumentPartRename(string baseFileName, DocumentSplitCriteria criteria)
        {
            _baseFileName = baseFileName;
            _criteria = criteria;
        }

        void IDocumentPartSavingCallback.DocumentPartSaving(DocumentPartSavingArgs args)
        {
            string partType = _criteria switch
            {
                DocumentSplitCriteria.SectionBreak => "Section",
                DocumentSplitCriteria.PageBreak => "Page",
                DocumentSplitCriteria.ColumnBreak => "Column",
                DocumentSplitCriteria.HeadingParagraph => "Heading",
                _ => "Part"
            };

            string newFileName = $"{Path.GetFileNameWithoutExtension(_baseFileName)} part {++_partIndex}, of type {partType}{Path.GetExtension(args.DocumentPartFileName)}";
            args.DocumentPartFileName = newFileName;
        }
    }
}
