using System;
using Aspose.Words;
using Aspose.Words.Drawing;

namespace DocumentJoinExample
{
    class Program
    {
        static void Main()
        {
            // Path to the folder that contains the source document and where the result will be saved.
            const string folderPath = @"C:\Docs\";

            // Load the document that will be inserted.
            Document srcDoc = new Document(folderPath + "Source.docx");

            // Create a new blank destination document.
            Document dstDoc = new Document();

            // Initialize a DocumentBuilder for the destination document.
            DocumentBuilder builder = new DocumentBuilder(dstDoc);

            // Add some initial content.
            builder.Writeln("This text appears before the inserted document.");

            // Insert a placeholder run that will serve as the insertion point.
            Run placeholderRun = new Run(dstDoc, "PLACEHOLDER");
            builder.CurrentParagraph.AppendChild(placeholderRun);

            // Move the builder's cursor to the placeholder run.
            builder.MoveTo(placeholderRun);

            // Insert the source document at the current cursor position.
            // KeepSourceFormatting preserves the original formatting of the inserted document.
            builder.InsertDocument(srcDoc, ImportFormatMode.KeepSourceFormatting);

            // Remove the placeholder run now that the insertion is complete.
            placeholderRun.Remove();

            // Add some trailing content (optional).
            builder.Writeln("This text appears after the inserted document.");

            // Save the combined document.
            dstDoc.Save(folderPath + "Result.docx");
        }
    }
}
