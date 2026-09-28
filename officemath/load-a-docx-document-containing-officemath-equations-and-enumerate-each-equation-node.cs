using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Fields;
using Aspose.Words.Math;

public class OfficeMathEnumerationExample
{
    public static void Main()
    {
        // Create a sample DOCX with a few OfficeMath equations using the EQ-field bootstrap workflow.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Sample equations:");

        string[] equations = new string[]
        {
            @"\f(1,2)",               // Simple fraction
            @"\r(2,x)",               // Square root of x
            @"\f(\r(2,x),\r(3,y))"    // Fraction with nested roots
        };

        foreach (string eq in equations)
        {
            // Insert a new paragraph for the equation.
            builder.Writeln();

            // Insert an EQ field.
            Field field = builder.InsertField(FieldType.FieldEquation, true);
            if (field is FieldEQ fieldEq)
            {
                // Write the EQ argument string into the field separator.
                builder.MoveTo(fieldEq.Separator);
                builder.Write(eq);

                // Convert the field to a real OfficeMath node.
                OfficeMath officeMath = fieldEq.AsOfficeMath();
                if (officeMath != null)
                {
                    // Insert the OfficeMath node before the field start.
                    Node fieldStart = fieldEq.Start;
                    if (fieldStart?.ParentNode is CompositeNode parent)
                    {
                        parent.InsertBefore(officeMath, fieldStart);
                    }

                    // Remove the original field.
                    fieldEq.Remove();
                }
            }
        }

        // Save the sample document.
        const string samplePath = "SampleEquations.docx";
        doc.Save(samplePath, SaveFormat.Docx);

        // Load the document and enumerate OfficeMath nodes.
        Document loadedDoc = new Document(samplePath);
        NodeCollection mathNodes = loadedDoc.GetChildNodes(NodeType.OfficeMath, true);

        StringBuilder reportBuilder = new StringBuilder();
        reportBuilder.AppendLine($"Total OfficeMath nodes: {mathNodes.Count}");
        int index = 0;
        foreach (OfficeMath om in mathNodes)
        {
            index++;
            reportBuilder.AppendLine($"Equation {index}:");
            reportBuilder.AppendLine($"  MathObjectType: {om.MathObjectType}");
            reportBuilder.AppendLine($"  DisplayType: {om.DisplayType}");
            reportBuilder.AppendLine($"  Justification: {om.Justification}");
            reportBuilder.AppendLine($"  Text: {om.GetText()}");
        }

        const string reportPath = "EquationReport.txt";
        File.WriteAllText(reportPath, reportBuilder.ToString());

        // Validate that the report file was created.
        if (!File.Exists(reportPath))
        {
            throw new Exception("Failed to create the equation report file.");
        }

        // Non‑interactive confirmation.
        Console.WriteLine("Enumeration completed. Report saved to " + reportPath);
    }
}
