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
        // Paths for output files
        string docPath = "OfficeMathSample.docx";
        string reportPath = "UnsupportedMathTypes.txt";

        // Create a new document and builder
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Helper to create an OfficeMath node from an EQ field
        void InsertEquation(string eqArgument)
        {
            // Insert an EQ field
            Field field = builder.InsertField(FieldType.FieldEquation, true);

            // Write the EQ argument into the field separator
            builder.MoveTo(field.Separator);
            builder.Write(eqArgument);

            // Convert the field to a real OfficeMath node
            OfficeMath officeMath = ((FieldEQ)field).AsOfficeMath();

            if (officeMath != null)
            {
                // Insert the OfficeMath node before the field start within the containing paragraph
                Paragraph paragraph = field.Start.ParentParagraph;
                if (paragraph != null)
                {
                    paragraph.InsertBefore(officeMath, field.Start);
                }

                // Remove the original field
                field.Remove();
            }
        }

        // Insert a simple fraction equation (will produce nested OfficeMath objects)
        builder.Writeln("Equation 1:");
        InsertEquation(@"\frac(1,2)");

        // Insert a function equation (e.g., sine)
        builder.Writeln();
        builder.Writeln("Equation 2:");
        InsertEquation(@"\sin(30)");

        // Save the document
        doc.Save(docPath, SaveFormat.Docx);

        // Reload the document to ensure a clean state
        Document loadedDoc = new Document(docPath);

        // Collect unsupported MathObjectTypes
        List<string> unsupportedLogs = new List<string>();
        NodeCollection officeMathNodes = loadedDoc.GetChildNodes(NodeType.OfficeMath, true);
        for (int i = 0; i < officeMathNodes.Count; i++)
        {
            OfficeMath om = (OfficeMath)officeMathNodes[i];
            MathObjectType type = om.MathObjectType;

            // Only OMathPara is considered supported for this example
            if (type != MathObjectType.OMathPara)
            {
                string log = $"Unsupported MathObjectType: {type} at node index {i}";
                unsupportedLogs.Add(log);
                Console.WriteLine(log);
            }
        }

        // Write the report to a text file
        File.WriteAllLines(reportPath, unsupportedLogs);

        // Validate that output files were created
        if (!File.Exists(docPath))
            throw new Exception($"Document was not saved to '{docPath}'.");
        if (!File.Exists(reportPath))
            throw new Exception($"Report was not saved to '{reportPath}'.");

        // Indicate completion
        Console.WriteLine("Processing completed.");
    }
}
