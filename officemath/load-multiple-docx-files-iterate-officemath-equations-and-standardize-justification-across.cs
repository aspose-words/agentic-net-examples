using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fields;
using Aspose.Words.Math;
using Aspose.Words.Saving;

public class OfficeMathJustificationBatch
{
    public static void Main()
    {
        // Folder for input and output documents.
        string inputFolder = "InputDocs";
        string outputFolder = "OutputDocs";
        Directory.CreateDirectory(inputFolder);
        Directory.CreateDirectory(outputFolder);

        // Define sample equations for each document.
        var docEquations = new Dictionary<string, string[]>
        {
            { Path.Combine(inputFolder, "Sample1.docx"), new[] { @"\f(1,2)", @"\r(3,x)" } },
            { Path.Combine(inputFolder, "Sample2.docx"), new[] { @"\i", @"\s \up5(Sup)", @"\b \bc\[ (\a \al \co2 \vs2 \hs2(1,0,0,0,1,0,0,0,1))" } }
        };

        // Create sample documents with equations.
        foreach (var kvp in docEquations)
        {
            CreateDocumentWithEquations(kvp.Key, kvp.Value);
        }

        // Process each document: standardize OfficeMath justification.
        foreach (string filePath in Directory.GetFiles(inputFolder, "*.docx"))
        {
            Document doc = new Document(filePath);

            // Get all OfficeMath nodes in the document.
            NodeCollection officeMathNodes = doc.GetChildNodes(NodeType.OfficeMath, true);

            foreach (OfficeMath officeMath in officeMathNodes)
            {
                // Only modify top‑level equations (MathObjectType.OMathPara).
                if (officeMath.MathObjectType == MathObjectType.OMathPara)
                {
                    // Ensure the equation is displayed on its own line before setting justification.
                    officeMath.DisplayType = OfficeMathDisplayType.Display;
                    officeMath.Justification = OfficeMathJustification.Center;
                }
            }

            // Save the modified document.
            string outputPath = Path.Combine(outputFolder, Path.GetFileNameWithoutExtension(filePath) + "_out.docx");
            doc.Save(outputPath, SaveFormat.Docx);
        }
    }

    // Creates a DOCX file containing the specified equations, each placed in its own paragraph.
    private static void CreateDocumentWithEquations(string filePath, string[] equations)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        foreach (string eqArgs in equations)
        {
            // Insert an EQ field.
            FieldEQ field = (FieldEQ)builder.InsertField(FieldType.FieldEquation, true);
            // Write the equation arguments into the field separator.
            builder.MoveTo(field.Separator);
            builder.Write(eqArgs);
            // Return the builder to the paragraph containing the field.
            builder.MoveTo(field.Start.ParentNode);

            // Convert the field to a real OfficeMath object.
            OfficeMath officeMath = field.AsOfficeMath();
            if (officeMath != null)
            {
                // Insert the OfficeMath node before the field start.
                field.Start.ParentNode.InsertBefore(officeMath, field.Start);
                // Remove the original field.
                field.Remove();
            }

            // Start a new paragraph for the next equation.
            builder.Writeln();
        }

        // Save the sample document.
        doc.Save(filePath, SaveFormat.Docx);
    }
}
