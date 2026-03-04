using System;
using Aspose.Words;

namespace DocumentInsertionExample
{
    class Program
    {
        static void Main()
        {
            // Load the destination document (the one that will receive the insertion).
            Document destination = new Document("Destination.docx");

            // Load the source document (the content to be inserted).
            Document source = new Document("Source.docx");

            // Create a DocumentBuilder for the destination document.
            DocumentBuilder builder = new DocumentBuilder(destination);

            // Move the builder's cursor to the beginning of the target section.
            // Sections are zero‑based; here we move to the second section (index 1).
            builder.MoveToSection(1);

            // Insert the source document at the current cursor position.
            // KeepSourceFormatting preserves the original styles of the inserted content.
            builder.InsertDocument(source, ImportFormatMode.KeepSourceFormatting);

            // Save the modified document.
            destination.Save("Result.docx");
        }
    }
}
