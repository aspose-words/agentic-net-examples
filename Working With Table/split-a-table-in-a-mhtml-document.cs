using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Words.Tables;

class SplitTableInMhtml
{
    static void Main()
    {
        // Load the source document that contains a table.
        Document doc = new Document("Input.docx");

        // Locate the first table in the document.
        Table table = doc.FirstSection.Body.Tables[0];

        // Insert a section break before the table.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.MoveTo(table);
        builder.InsertBreak(BreakType.SectionBreakContinuous);

        // Insert a section break after the table.
        // After the previous insertion the cursor is positioned before the table, so we move
        // to the node that follows the table and insert the break there.
        Node nextNode = table.NextSibling;
        if (nextNode != null)
        {
            builder.MoveTo(nextNode);
            builder.InsertBreak(BreakType.SectionBreakContinuous);
        }
        else
        {
            // If the table is the last node in the body, move to the end of the body.
            builder.MoveTo(doc.LastSection.Body);
            builder.Writeln(); // ensure we are after the table
            builder.InsertBreak(BreakType.SectionBreakContinuous);
        }

        // Configure save options for MHTML with document splitting at section breaks.
        HtmlSaveOptions saveOptions = new HtmlSaveOptions(SaveFormat.Mhtml)
        {
            DocumentSplitCriteria = DocumentSplitCriteria.SectionBreak,
            // Optional: give each part a meaningful name.
            DocumentPartSavingCallback = new SavedDocumentPartRename("SplitTable")
        };

        // Save the document as MHTML. Each section (including the table) will be saved as a separate part.
        doc.Save("Output.mhtml", saveOptions);
    }

    // Callback to customize the filenames of the split document parts.
    private class SavedDocumentPartRename : IDocumentPartSavingCallback
    {
        private readonly string _baseName;
        private int _partIndex = 0;

        public SavedDocumentPartRename(string baseName)
        {
            _baseName = baseName;
        }

        void IDocumentPartSavingCallback.DocumentPartSaving(DocumentPartSavingArgs args)
        {
            // Generate a filename like "SplitTable part 1.mhtml", "SplitTable part 2.mhtml", etc.
            string extension = Path.GetExtension(args.DocumentPartFileName);
            string partFileName = $"{_baseName} part {++_partIndex}{extension}";
            args.DocumentPartFileName = partFileName;
        }
    }
}
