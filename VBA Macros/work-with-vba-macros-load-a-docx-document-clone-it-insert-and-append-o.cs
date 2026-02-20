using System;
using Aspose.Words;
using Aspose.Words.Vba;

class VbaDocumentProcessor
{
    static void Main()
    {
        // Load the original DOCX (macro‑enabled) document.
        Document originalDoc = new Document("Original.docx");

        // Clone the whole document, including its VBA project.
        Document clonedDoc = (Document)originalDoc.Clone();
        if (originalDoc.HasMacros)
        {
            // Clone the VBA project and assign it to the cloned document.
            VbaProject clonedVba = originalDoc.VbaProject.Clone();
            clonedDoc.VbaProject = clonedVba;
        }

        // Load documents that will be inserted and appended.
        Document docToInsert = new Document("Insert.docx");
        Document docToAppend = new Document("Append.docx");

        // Insert the document at the beginning of the cloned document.
        DocumentBuilder builder = new DocumentBuilder(clonedDoc);
        builder.MoveToDocumentStart();
        builder.InsertDocument(docToInsert, ImportFormatMode.KeepSourceFormatting);

        // Append the document at the end of the cloned document.
        clonedDoc.AppendDocument(docToAppend, ImportFormatMode.KeepSourceFormatting);

        // Split the resulting document into separate files, one per section.
        for (int i = 0; i < clonedDoc.Sections.Count; i++)
        {
            // Create a new empty document.
            Document splitDoc = new Document();

            // Import the current section into the new document.
            NodeImporter importer = new NodeImporter(clonedDoc, splitDoc, ImportFormatMode.KeepSourceFormatting);
            Section importedSection = (Section)importer.ImportNode(clonedDoc.Sections[i], true);
            splitDoc.Sections.Clear(); // Remove the default empty section.
            splitDoc.Sections.Add(importedSection);

            // Preserve the VBA project in each split document if needed.
            if (clonedDoc.HasMacros)
            {
                splitDoc.VbaProject = clonedDoc.VbaProject.Clone();
            }

            // Save the split document.
            string outFileName = $"Split_Part_{i + 1}.docx";
            splitDoc.Save(outFileName);
        }

        // Finally, save the combined cloned document.
        clonedDoc.Save("CombinedResult.docx");
    }
}
