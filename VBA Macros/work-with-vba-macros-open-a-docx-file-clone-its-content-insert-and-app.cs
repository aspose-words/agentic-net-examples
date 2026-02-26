using System;
using Aspose.Words;
using Aspose.Words.Saving;

class VbaMacroDocumentProcessor
{
    static void Main()
    {
        // Path to the folder that contains the source documents.
        string dataDir = @"C:\Docs\"; // <-- adjust to your environment

        // 1. Load the original DOCX (or DOCM) file.
        Document originalDoc = new Document(dataDir + "Source.docx");

        // Optional: check whether the document contains VBA macros.
        if (originalDoc.HasMacros)
        {
            Console.WriteLine("The source document contains VBA macros.");
        }
        else
        {
            Console.WriteLine("The source document does not contain VBA macros.");
        }

        // 2. Clone the original document (deep copy, including all nodes).
        Document clonedDoc = (Document)originalDoc.Clone(true);

        // 3. Insert an additional document into the cloned document at the end.
        Document docToInsert = new Document(dataDir + "Insert.docx");
        DocumentBuilder builder = new DocumentBuilder(clonedDoc);
        builder.MoveToDocumentEnd(); // Position the cursor at the end of the cloned document.
        builder.InsertDocument(docToInsert, ImportFormatMode.KeepSourceFormatting);

        // 4. Append another document to the cloned document.
        Document docToAppend = new Document(dataDir + "Append.docx");
        clonedDoc.AppendDocument(docToAppend, ImportFormatMode.KeepSourceFormatting);

        // 5. Split the original document into two parts:
        //    - Part 1: pages 1 to 2
        //    - Part 2: remaining pages
        int totalPages = originalDoc.PageCount;
        Document part1 = originalDoc.ExtractPages(1, Math.Min(2, totalPages));
        Document part2 = originalDoc.ExtractPages(3, totalPages - 2);

        // 6. Save all resulting documents.
        originalDoc.Save(dataDir + "Original_Saved.docx");
        clonedDoc.Save(dataDir + "Cloned_Modified.docx");
        part1.Save(dataDir + "Original_Part1.docx");
        part2.Save(dataDir + "Original_Part2.docx");

        Console.WriteLine("Processing completed.");
    }
}
