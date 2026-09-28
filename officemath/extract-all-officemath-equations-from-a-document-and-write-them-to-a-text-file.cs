using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fields;
using Aspose.Words.Math;

public class Program
{
    public static void Main()
    {
        // Create a new document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert sample equations using the deterministic EQ‑field bootstrap workflow.
        InsertEquation(builder, @"\f(1,2)"); // Simple fraction.
        InsertEquation(builder, @"\r(3,x)"); // Simple root.

        // Save the sample document (optional, demonstrates that the document contains the equations).
        string docPath = "Sample.docx";
        doc.Save(docPath);

        // Extract all top‑level OfficeMath equations.
        NodeCollection officeMathNodes = doc.GetChildNodes(NodeType.OfficeMath, true);
        List<string> reportLines = new List<string>();
        int equationIndex = 1;

        foreach (OfficeMath om in officeMathNodes)
        {
            // Consider only top‑level equations (OMathPara).
            if (om.MathObjectType == MathObjectType.OMathPara)
            {
                string equationText = om.GetText().Trim();
                reportLines.Add($"Equation {equationIndex}: {equationText}");
                equationIndex++;
            }
        }

        // Write the extracted equations to a text file.
        string reportPath = "EquationsReport.txt";
        File.WriteAllLines(reportPath, reportLines);

        // Validate that the report file was created.
        if (!File.Exists(reportPath))
        {
            throw new Exception("The equations report file was not created.");
        }

        Console.WriteLine($"Extraction complete. Report saved to '{reportPath}'.");
    }

    private static void InsertEquation(DocumentBuilder builder, string eqArgument)
    {
        // Insert an EQ field.
        FieldEQ field = builder.InsertField(FieldType.FieldEquation, true) as FieldEQ;
        if (field == null)
            return;

        // Write the EQ argument into the field separator.
        builder.MoveTo(field.Separator);
        builder.Write(eqArgument);

        // Convert the field to a real OfficeMath object.
        OfficeMath officeMath = field.AsOfficeMath();
        if (officeMath != null)
        {
            // Insert the OfficeMath node before the field start.
            Node startNode = field.Start;
            startNode.ParentNode.InsertBefore(officeMath, startNode);

            // Remove the original field.
            field.Remove();
        }
    }
}
