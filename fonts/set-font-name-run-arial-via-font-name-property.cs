using System;
using Aspose.Words;

class Program
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Get the first paragraph (created by default) to host the run.
        Paragraph paragraph = doc.FirstSection.Body.FirstParagraph;

        // Create a run with some text.
        Run run = new Run(doc, "Hello world!");

        // Set the font name of the run to Arial.
        run.Font.Name = "Arial";

        // Append the run to the paragraph.
        paragraph.AppendChild(run);

        // Save the document to a file.
        doc.Save("RunWithArial.docx");
    }
}
