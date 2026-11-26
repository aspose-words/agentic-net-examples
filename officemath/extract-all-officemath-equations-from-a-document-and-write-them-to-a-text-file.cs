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
        // Create a sample document with a few equations.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        InsertEquation(builder, @"\f(1,2)");
        builder.Writeln();

        InsertEquation(builder, @"\r(3,x)");
        builder.Writeln();

        InsertEquation(builder, @"\s(5)");
        builder.Writeln();

        // Save the sample document.
        string docPath = "Sample.docx";
        doc.Save(docPath, SaveFormat.Docx);

        // Load the document (could reuse the same instance).
        Document loadedDoc = new Document(docPath);

        // Extract top‑level OfficeMath equations.
        NodeCollection mathNodes = loadedDoc.GetChildNodes(NodeType.OfficeMath, true);
        List<string> lines = new List<string>();
        int equationIndex = 1;

        foreach (OfficeMath om in mathNodes)
        {
            if (om.MathObjectType == MathObjectType.OMathPara)
            {
                string text = om.GetText().Trim();
                lines.Add($"Equation {equationIndex}: {text}");
                equationIndex++;
            }
        }

        // Write the extracted equations to a text file.
        string reportPath = "Equations.txt";
        File.WriteAllLines(reportPath, lines);

        // Validate that the report file was created.
        if (!File.Exists(reportPath))
        {
            throw new Exception("Failed to create the equations report file.");
        }

        Console.WriteLine($"Extracted {equationIndex - 1} equations to \"{reportPath}\".");
    }

    private static void InsertEquation(DocumentBuilder builder, string eqSwitch)
    {
        // Insert an Equation field.
        Field field = builder.InsertField(FieldType.FieldEquation, true);
        FieldEQ fieldEq = field as FieldEQ;

        if (fieldEq != null)
        {
            // Write the EQ argument into the field separator.
            builder.MoveTo(fieldEq.Separator);
            builder.Write(eqSwitch);

            // Convert the field to a real OfficeMath node.
            OfficeMath officeMath = fieldEq.AsOfficeMath();
            if (officeMath != null)
            {
                // Insert the OfficeMath node before the field start.
                Node startNode = fieldEq.Start;
                startNode.ParentNode.InsertBefore(officeMath, startNode);
                // Remove the original field.
                fieldEq.Remove();
            }
        }
    }
}
