using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fields;
using Aspose.Words.Math;

public class Program
{
    private const string OutputDocPath = "Equations.docx";
    private const string ReportPath = "EquationCount.txt";

    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert several equations using the deterministic EQ‑field bootstrap workflow.
        InsertEquation(builder, @"\f(1,2)");
        InsertEquation(builder, @"\r(3,x)");
        InsertEquation(builder, @"\s(5,7)");

        // Save the document.
        doc.Save(OutputDocPath);

        // Verify the document was saved.
        if (!File.Exists(OutputDocPath))
            throw new Exception($"Failed to create output document '{OutputDocPath}'.");

        // Count all top‑level OfficeMath nodes (equations).
        NodeCollection officeMathNodes = doc.GetChildNodes(NodeType.OfficeMath, true);
        int equationCount = 0;
        foreach (OfficeMath om in officeMathNodes)
        {
            if (om.MathObjectType == MathObjectType.OMathPara)
                equationCount++;
        }

        // Write the count to a report file.
        File.WriteAllText(ReportPath, equationCount.ToString());

        // Verify the report file was created.
        if (!File.Exists(ReportPath))
            throw new Exception($"Failed to create report file '{ReportPath}'.");

        // Output the result to the console.
        Console.WriteLine($"Total equations (top‑level OfficeMath nodes): {equationCount}");
    }

    private static void InsertEquation(DocumentBuilder builder, string eqSwitch)
    {
        // Insert an EQ field.
        Field field = builder.InsertField(FieldType.FieldEquation, true);
        if (field is not FieldEQ fieldEQ)
            throw new Exception("Inserted field is not a FieldEQ.");

        // Write the EQ argument.
        builder.MoveTo(fieldEQ.Separator);
        builder.Write(eqSwitch);

        // Update the field so that Aspose.Words can convert it to OfficeMath.
        field.Update();

        // Convert the field to an OfficeMath node.
        OfficeMath? officeMath = fieldEQ.AsOfficeMath();
        if (officeMath == null)
            throw new Exception("Failed to convert EQ field to OfficeMath.");

        // Insert the OfficeMath node before the field start.
        Node? fieldStart = field.Start;
        if (fieldStart == null)
            throw new Exception("Field start node is null.");
        fieldStart.ParentNode!.InsertBefore(officeMath, fieldStart);

        // Remove the original field.
        field.Remove();

        // Move the builder to the end of the document for the next insertion.
        builder.MoveToDocumentEnd();
    }
}
