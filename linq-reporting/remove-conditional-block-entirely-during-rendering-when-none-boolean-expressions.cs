using System;
using System.Data;
using Aspose.Words;
using Aspose.Words.Fields;

class RemoveFalseIfBlocks
{
    static void Main()
    {
        // Load the source document.
        Document doc = new Document("Input.docx");

        // Iterate over all IF fields in the document.
        foreach (FieldIf ifField in doc.Range.Fields)
        {
            // Evaluate the condition of the IF field.
            FieldIfComparisonResult result = ifField.EvaluateCondition();

            // If the condition is false, remove the entire IF block (field code and displayed result).
            if (result == FieldIfComparisonResult.False)
            {
                // Capture the start and end nodes of the field.
                Node startNode = ifField.Start;
                Node endNode = ifField.End;

                // Remove all nodes between the start and end (including the result text).
                Node curNode = startNode;
                while (curNode != null && curNode != endNode)
                {
                    Node nextNode = curNode.NextSibling;
                    curNode.Remove();
                    curNode = nextNode;
                }

                // Finally remove the end node itself.
                if (endNode != null)
                    endNode.Remove();
            }
        }

        // Save the modified document.
        doc.Save("Output.docx");
    }
}
