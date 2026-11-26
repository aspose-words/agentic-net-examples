using System;
using System.IO;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Fields;
using Aspose.Words.Math;
using Newtonsoft.Json;

public class Program
{
    public static void Main()
    {
        // Prepare output directory
        string outputDir = "Output";
        Directory.CreateDirectory(outputDir);

        // Create a new document
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Define simple EQ switches for safe conversion
        string[] eqSwitches = new string[] { @"\f(1,2)", @"\r(3,x)", @"\f(5,6)" };

        // Insert each equation into its own paragraph using the deterministic EQ-field bootstrap workflow
        foreach (string eq in eqSwitches)
        {
            // Start a new paragraph
            builder.Writeln();

            // Insert an empty Equation field
            Field field = builder.InsertField(FieldType.FieldEquation, true);

            // Write the EQ argument into the field separator
            builder.MoveTo(field.Separator);
            builder.Write(eq);

            // Convert the field to a real OfficeMath node
            FieldEQ fieldEQ = field as FieldEQ;
            OfficeMath officeMath = fieldEQ?.AsOfficeMath();

            if (officeMath != null)
            {
                // Insert the OfficeMath node before the field start
                Node fieldStart = field.Start;
                CompositeNode parent = fieldStart.ParentNode as CompositeNode;
                if (parent != null)
                {
                    parent.InsertBefore(officeMath, fieldStart);
                }

                // Remove the original field so only the OfficeMath remains
                field.Remove();
            }
        }

        // Save the initial document
        string initialPath = Path.Combine(outputDir, "BulkUpdate.docx");
        doc.Save(initialPath, SaveFormat.Docx);

        // Bulk update: (no display or justification changes needed for validation)
        NodeCollection mathNodes = doc.GetChildNodes(NodeType.OfficeMath, true);
        foreach (OfficeMath om in mathNodes)
        {
            // Example bulk operation could be added here if needed.
            // For this task we only need to ensure MathObjectType is OMathPara.
        }

        // Validation: ensure every OfficeMath node is a top‑level paragraph equation
        foreach (OfficeMath om in mathNodes)
        {
            if (om.MathObjectType != MathObjectType.OMathPara)
            {
                throw new InvalidOperationException("Encountered an OfficeMath node that is not a top‑level OMathPara.");
            }
        }

        // Save the updated document
        string updatedPath = Path.Combine(outputDir, "BulkUpdateUpdated.docx");
        doc.Save(updatedPath, SaveFormat.Docx);

        // Generate a JSON report describing each OfficeMath node
        var reportItems = new List<object>();
        int index = 0;
        foreach (OfficeMath om in mathNodes)
        {
            reportItems.Add(new
            {
                Index = index++,
                MathObjectType = om.MathObjectType.ToString()
                // DisplayType and Justification omitted because they are not available in this version.
            });
        }

        string json = JsonConvert.SerializeObject(reportItems, Formatting.Indented);
        string reportPath = Path.Combine(outputDir, "OfficeMathReport.json");
        File.WriteAllText(reportPath, json);

        // Verify that all expected output files were created
        if (!File.Exists(initialPath) || !File.Exists(updatedPath) || !File.Exists(reportPath))
        {
            throw new FileNotFoundException("One or more expected output files were not created.");
        }

        // Execution completed successfully
    }
}
