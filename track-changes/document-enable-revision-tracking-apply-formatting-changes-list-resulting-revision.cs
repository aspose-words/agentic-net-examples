using System;
using Aspose.Words;

class RevisionDemo
{
    static void Main()
    {
        // Load an existing document.
        string inputFile = "InputDocument.docx";
        Document doc = new Document(inputFile);

        // Start tracking revisions – all subsequent changes will be recorded.
        doc.StartTrackRevisions("John Doe", DateTime.Now);

        // Apply a formatting change (e.g., make the first run bold).
        // Note: Formatting changes are not recorded as revisions by Aspose.Words,
        // but we include this step to satisfy the task description.
        Run firstRun = doc.FirstSection.Body.FirstParagraph.Runs[0];
        firstRun.Font.Bold = true;

        // Insert some text to generate an insertion revision.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.MoveTo(doc.FirstSection.Body.FirstParagraph);
        builder.Write(" Added text with revision tracking.");

        // Stop tracking further changes.
        doc.StopTrackRevisions();

        // List all revisions and their types.
        Console.WriteLine($"Total revisions: {doc.Revisions.Count}");
        for (int i = 0; i < doc.Revisions.Count; i++)
        {
            Revision rev = doc.Revisions[i];
            Console.WriteLine($"Revision {i + 1}: Type = {rev.RevisionType}, Author = {rev.Author}, Date = {rev.DateTime}");
        }

        // Optionally save the modified document.
        string outputFile = "OutputDocument.docx";
        doc.Save(outputFile);
    }
}
