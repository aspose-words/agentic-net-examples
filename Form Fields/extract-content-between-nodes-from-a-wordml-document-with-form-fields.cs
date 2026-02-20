using System;
using Aspose.Words;
using Aspose.Words.Fields;

class ExtractFormFieldContent
{
    static void Main()
    {
        // Load the WORDML (or any supported) document.
        Document doc = new Document("input.docx");

        // Iterate through all fields in the document.
        foreach (Field field in doc.Range.Fields)
        {
            // We are interested only in FORMTEXT fields (form fields that contain user‑editable text).
            if (field is FieldFormText formTextField)
            {
                // The text entered by the user is stored in the Result property.
                // DisplayResult returns the same text as it appears in the document.
                string fieldResult = formTextField.Result;
                Console.WriteLine($"Form field result: {fieldResult}");

                // If you need the exact range of nodes that belong to this field,
                // you can obtain the start and end nodes and create a sub‑range.
                Node startNode = formTextField.Start;
                Node endNode = formTextField.End;

                // Build a range that spans from the start node (inclusive) to the end node (exclusive).
                // The Range constructor is not public, so we use the Document's Range and then
                // extract the needed text via GetText on the nodes between start and end.
                Node current = startNode;
                while (current != null && current != endNode)
                {
                    // Process each node inside the field if needed.
                    // For demonstration we just output its text.
                    Console.WriteLine($"  Node text: {current.GetText()}");
                    current = current.NextSibling;
                }
            }
        }
    }
}
