using System;
using Aspose.Words;
using Aspose.Words.Markup;

namespace ContentControlValidation
{
    class Program
    {
        static void Main()
        {
            // Load the document (lifecycle rule: load)
            Document doc = new Document("Input.docx");

            // Ensure the document has the minimal required nodes (lifecycle rule: create/ensure)
            doc.EnsureMinimum();

            // Validate that all required Structured Document Tags (content controls) contain non‑empty text.
            ValidateContentControls(doc);

            // Save the document (lifecycle rule: save)
            doc.Save("ValidatedOutput.docx");
        }

        /// <summary>
        /// Throws an exception if any StructuredDocumentTag in the document has empty or whitespace text.
        /// </summary>
        /// <param name="doc">The document to validate.</param>
        private static void ValidateContentControls(Document doc)
        {
            // Retrieve all content controls (StructuredDocumentTag nodes) in the document.
            NodeCollection tags = doc.GetChildNodes(NodeType.StructuredDocumentTag, true);

            foreach (Node node in tags)
            {
                StructuredDocumentTag tag = (StructuredDocumentTag)node;

                // Get the plain text inside the content control.
                // The GetText method returns the text of the tag and its children.
                string innerText = tag.GetText();

                // Trim whitespace to detect truly empty content.
                if (string.IsNullOrWhiteSpace(innerText))
                {
                    // You can customize the message – here we include the tag's Id for easier debugging.
                    throw new InvalidOperationException(
                        $"Content control with Id={tag.Id} is required but contains no text.");
                }
            }
        }
    }
}
