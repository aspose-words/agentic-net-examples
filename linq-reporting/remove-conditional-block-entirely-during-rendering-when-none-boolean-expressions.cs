using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Fields;

class RemoveFalseIfBlocks
{
    static void Main()
    {
        // Load the source document.
        Document doc = new Document("Input.docx");

        // Collect all IF fields in the document.
        List<FieldIf> ifFields = new List<FieldIf>();
        foreach (Field field in doc.Range.Fields)
        {
            if (field is FieldIf ifField)
                ifFields.Add(ifField);
        }

        // Iterate over the collected IF fields.
        foreach (FieldIf ifField in ifFields)
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
                endNode?.Remove();
            }
        }

        // Save the modified document.
        doc.Save("Output.docx");
    }
}
