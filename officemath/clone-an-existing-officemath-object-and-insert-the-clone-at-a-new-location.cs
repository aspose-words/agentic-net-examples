using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fields;
using Aspose.Words.Math;

public class CloneOfficeMathExample
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add a paragraph with some introductory text.
        builder.Writeln("Original equation:");

        // Insert an EQ field that will be converted to a real OfficeMath object.
        Field field = builder.InsertField(FieldType.FieldEquation, true);
        if (field is not FieldEQ fieldEq)
            throw new InvalidOperationException("Inserted field is not a FieldEQ.");

        // Write a simple equation into the field separator.
        builder.MoveTo(fieldEq.Separator);
        builder.Write(@"\f(1,2)"); // fraction 1 over 2

        // Update the field so that the equation is parsed.
        field.Update();

        // Convert the field to an OfficeMath node.
        OfficeMath originalMath = fieldEq.AsOfficeMath();
        if (originalMath == null)
            throw new InvalidOperationException("Failed to convert field to OfficeMath.");

        // Insert the OfficeMath node before the field start and remove the field.
        Node fieldStart = fieldEq.Start;
        fieldStart.ParentNode.InsertBefore(originalMath, fieldStart);
        fieldEq.Remove();

        // Clone the existing OfficeMath node.
        OfficeMath clonedMath = (OfficeMath)originalMath.Clone(true);

        // Insert a new paragraph and place the cloned OfficeMath there.
        builder.Writeln(); // creates a new empty paragraph.
        builder.InsertNode(clonedMath);

        // Save the document.
        string outputPath = "CloneOfficeMath.docx";
        doc.Save(outputPath);

        // Validate that the file was saved and contains two top‑level OfficeMath nodes.
        if (!File.Exists(outputPath))
            throw new FileNotFoundException("Output file was not created.", outputPath);

        Document loadedDoc = new Document(outputPath);
        NodeCollection mathNodes = loadedDoc.GetChildNodes(NodeType.OfficeMath, true);
        int topLevelCount = 0;
        foreach (OfficeMath om in mathNodes)
        {
            if (om.MathObjectType == MathObjectType.OMathPara)
                topLevelCount++;
        }

        if (topLevelCount != 2)
            throw new Exception($"Expected 2 top‑level OfficeMath nodes, but found {topLevelCount}.");

        Console.WriteLine($"Document saved to '{outputPath}' and contains {topLevelCount} equations.");
    }
}
