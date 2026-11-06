using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fields;
using Aspose.Words.Math;

public class Program
{
    public static void Main()
    {
        // Create a new document and builder.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert an EQ field that will be converted to a real OfficeMath node.
        Field field = builder.InsertField(FieldType.FieldEquation, true);
        FieldEQ fieldEQ = (FieldEQ)field;

        // Write a simple equation into the field separator.
        builder.MoveTo(fieldEQ.Separator);
        builder.Write("\\f(1,2)"); // fraction 1/2

        // Update the field so that Aspose.Words evaluates it.
        field.Update();

        // Convert the field to OfficeMath.
        OfficeMath officeMath = fieldEQ.AsOfficeMath();

        if (officeMath == null)
            throw new InvalidOperationException("Failed to convert EQ field to OfficeMath.");

        // Insert the OfficeMath node before the field start.
        Node fieldStart = field.Start;
        fieldStart.ParentNode.InsertBefore(officeMath, fieldStart);

        // Remove the original field (start, separator, end).
        field.Remove();

        // Save the document to verify output.
        string outputPath = "Output.docx";
        doc.Save(outputPath);
        if (!File.Exists(outputPath))
            throw new FileNotFoundException("The output document was not created.", outputPath);

        // Enumerate OfficeMath nodes and test the matching function.
        NodeCollection mathNodes = doc.GetChildNodes(NodeType.OfficeMath, true);
        for (int i = 0; i < mathNodes.Count; i++)
        {
            OfficeMath om = (OfficeMath)mathNodes[i];
            bool isPara = IsMatchingMathObjectType(om, MathObjectType.OMathPara);
            Console.WriteLine($"OfficeMath node at collection index {i} matches OMathPara: {isPara}");
        }
    }

    // Returns true if the OfficeMath node's MathObjectType equals the specified target.
    public static bool IsMatchingMathObjectType(OfficeMath om, MathObjectType target)
    {
        if (om == null)
            return false;

        return om.MathObjectType == target;
    }
}
