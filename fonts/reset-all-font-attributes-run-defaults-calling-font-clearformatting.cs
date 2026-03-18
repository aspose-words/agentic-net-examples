using System;
using Aspose.Words;

class Program
{
    static void Main()
    {
        // Load an existing document.
        Document doc = new Document("Input.docx");

        // Iterate over every Run in the document.
        foreach (Run run in doc.GetChildNodes(NodeType.Run, true))
        {
            // Reset all font attributes of the run to their default values.
            run.Font.ClearFormatting();
        }

        // Save the document with the cleared font formatting.
        doc.Save("Output.docx");
    }
}
