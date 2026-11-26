using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fields;
using Aspose.Words.Math;

public class OfficeMathInspector
{
    public static void Main()
    {
        // Create a new document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Helper to create a safe OfficeMath equation via EQ field bootstrap.
        void InsertEquation(string eqSwitch)
        {
            // Insert an empty EQ field.
            Field field = builder.InsertField(FieldType.FieldEquation, true);
            // Write the equation switch into the field separator.
            builder.MoveTo(field.Separator);
            builder.Write(eqSwitch);
            // Convert the field to a real OfficeMath node.
            OfficeMath? om = ((FieldEQ)field).AsOfficeMath();
            if (om != null)
            {
                // Insert the OfficeMath node before the field start.
                field.Start.ParentNode.InsertBefore(om, field.Start);
                // Remove the original field.
                field.Remove();
            }
        }

        // Insert several sample equations.
        InsertEquation(@"\f(1,2)");      // Fraction
        builder.Writeln(); // New paragraph
        InsertEquation(@"\r(3,x)");      // Root
        builder.Writeln(); // New paragraph
        InsertEquation(@"\int x dx");   // Integral (simple)
        builder.Writeln(); // New paragraph

        // Save the document.
        string docPath = "Sample.docx";
        doc.Save(docPath, SaveFormat.Docx);

        // Reload the document to ensure a clean state.
        Document loadedDoc = new Document(docPath);

        // Collect unsupported MathObjectTypes (those not OMathPara).
        List<string> unsupportedTypes = new List<string>();
        NodeCollection mathNodes = loadedDoc.GetChildNodes(NodeType.OfficeMath, true);
        foreach (OfficeMath om in mathNodes)
        {
            if (om.MathObjectType != MathObjectType.OMathPara)
            {
                unsupportedTypes.Add(om.MathObjectType.ToString());
            }
        }

        // Prepare report content.
        string reportPath = "UnsupportedEquations.txt";
        using (StreamWriter writer = new StreamWriter(reportPath))
        {
            if (unsupportedTypes.Count == 0)
            {
                writer.WriteLine("No unsupported equation types found.");
                Console.WriteLine("No unsupported equation types found.");
            }
            else
            {
                writer.WriteLine("Unsupported MathObjectTypes detected:");
                foreach (string typeName in unsupportedTypes)
                {
                    writer.WriteLine("- " + typeName);
                }
                Console.WriteLine("Unsupported MathObjectTypes detected:");
                foreach (string typeName in unsupportedTypes)
                {
                    Console.WriteLine("- " + typeName);
                }
            }
        }

        // Validate that the report file was created.
        if (!File.Exists(reportPath))
        {
            throw new InvalidOperationException("Report file was not created.");
        }
    }
}
