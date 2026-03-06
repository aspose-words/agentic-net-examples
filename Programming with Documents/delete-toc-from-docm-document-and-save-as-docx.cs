using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Fields;
using Aspose.Words.Saving;

class DeleteTocExample
{
    static void Main()
    {
        // Input DOCM file path
        string inputPath = "input.docm";

        // Output DOCX file path
        string outputPath = "output.docx";

        // Load the macro-enabled document (DOCM)
        Document doc = new Document(inputPath);

        // Collect all Table of Contents (TOC) fields in the document
        List<Field> tocFields = new List<Field>();
        foreach (Field field in doc.Range.Fields)
        {
            if (field.Type == FieldType.FieldTOC)
                tocFields.Add(field);
        }

        // Remove each TOC field together with its result text
        foreach (Field toc in tocFields)
        {
            // The field consists of a start node, separator (optional), and end node.
            // Remove everything from the start node up to and including the end node.
            Node startNode = toc.Start;
            Node endNode = toc.End;

            Node current = startNode;
            while (current != null)
            {
                Node next = current.NextSibling;
                current.Remove(); // Detach the node from its parent
                if (current == endNode)
                    break;
                current = next;
            }
        }

        // Save the modified document as DOCX (macro-free)
        doc.Save(outputPath, SaveFormat.Docx);
    }
}
