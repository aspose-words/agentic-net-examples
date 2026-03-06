using System;
using Aspose.Words;

class Program
{
    static void Main()
    {
        // Paths to the documents.
        string destinationPath = "Destination.docx";
        string sourcePath = "Source.docx";
        string resultPath = "Result.docx";

        // Load the destination document (create rule: Document(string)).
        Document destinationDoc = new Document(destinationPath);

        // Create a DocumentBuilder for the destination document.
        DocumentBuilder builder = new DocumentBuilder(destinationDoc);

        // Locate the paragraph that will serve as the insertion point.
        // Here we use the first paragraph in the body of the first section.
        Paragraph insertionParagraph = destinationDoc.FirstSection.Body.FirstParagraph;

        // Move the builder's cursor to the chosen paragraph (lifecycle rule).
        builder.MoveTo(insertionParagraph);

        // Load the source document that will be inserted.
        Document sourceDoc = new Document(sourcePath);

        // Insert the source document at the current cursor position.
        // KeepSourceFormatting preserves the original formatting of the source.
        builder.InsertDocument(sourceDoc, ImportFormatMode.KeepSourceFormatting);

        // Save the resulting document (save rule).
        destinationDoc.Save(resultPath);
    }
}
