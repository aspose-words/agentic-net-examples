using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fields;
using Aspose.Words.Math;

public class Program
{
    // Returns true if the OfficeMath node's MathObjectType equals the specified target type.
    public static bool IsMathObjectTypeMatch(OfficeMath officeMath, MathObjectType targetType)
    {
        if (officeMath == null)
            return false;

        return officeMath.MathObjectType == targetType;
    }

    // Bootstrap helper: creates a top‑level OfficeMath node from a simple EQ string.
    private static OfficeMath CreateOfficeMath(Document doc, string eqSwitch)
    {
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert an EQ field.
        Field field = builder.InsertField(FieldType.FieldEquation, true);
        // Write the EQ switch into the field separator.
        builder.MoveTo(field.Separator);
        builder.Write(eqSwitch);

        // Update the field so that Aspose.Words parses it.
        field.Update();

        // Convert the field to OfficeMath.
        FieldEQ fieldEq = field as FieldEQ;
        OfficeMath officeMath = fieldEq?.AsOfficeMath();

        if (officeMath == null)
            throw new InvalidOperationException("Failed to convert EQ field to OfficeMath.");

        // Insert the OfficeMath node before the field start and remove the field.
        CompositeNode parent = field.Start.ParentNode as CompositeNode;
        if (parent == null)
            throw new InvalidOperationException("Unable to locate a valid parent for insertion.");

        parent.InsertBefore(officeMath, field.Start);
        field.Remove();

        return officeMath;
    }

    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Bootstrap a simple fraction equation (1/2) as a top‑level OfficeMath node.
        OfficeMath officeMath = CreateOfficeMath(doc, @"\f(1,2)");

        // Save the document.
        string outputPath = "Output.docx";
        doc.Save(outputPath);

        // Validate that the file was created.
        if (!File.Exists(outputPath))
            throw new FileNotFoundException("The output document was not saved.", outputPath);

        // Reload the document to ensure persistence.
        Document loadedDoc = new Document(outputPath);

        // Get all OfficeMath nodes.
        NodeCollection officeMathNodes = loadedDoc.GetChildNodes(NodeType.OfficeMath, true);

        // Count how many top‑level equations (MathObjectType.OMathPara) exist.
        int matchingCount = 0;
        foreach (OfficeMath om in officeMathNodes)
        {
            if (IsMathObjectTypeMatch(om, MathObjectType.OMathPara))
                matchingCount++;
        }

        // Output the result.
        Console.WriteLine($"Total OfficeMath nodes: {officeMathNodes.Count}");
        Console.WriteLine($"Top‑level OMathPara nodes matching criteria: {matchingCount}");
    }
}
