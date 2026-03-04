using System;
using Aspose.Words;

namespace InsertDocumentUsingRun
{
    class Program
    {
        static void Main()
        {
            // Create a new blank destination document.
            Document dstDoc = new Document();

            // Get the first paragraph of the destination document (it always exists in a new Document).
            Paragraph dstParagraph = dstDoc.FirstSection.Body.FirstParagraph;

            // Load the source DOCX file whose contents we want to insert.
            Document srcDoc = new Document("Source.docx");

            // Extract the plain text from the source document.
            string sourceText = srcDoc.GetText();

            // Create a Run in the destination document containing the source text.
            Run run = new Run(dstDoc, sourceText);

            // Append the Run to the destination paragraph.
            dstParagraph.AppendChild(run);

            // Save the resulting document.
            dstDoc.Save("Result.docx");
        }
    }
}
