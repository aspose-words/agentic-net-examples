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
        // Prepare directories.
        string dataDir = Path.Combine(Directory.GetCurrentDirectory(), "Data");
        Directory.CreateDirectory(dataDir);

        // Path for the sample document.
        string docPath = Path.Combine(dataDir, "SampleWithEquations.docx");

        // Create a document and insert a few equations using the deterministic EQ‑field bootstrap workflow.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.Writeln("Equation 1:");
        InsertOfficeMathEquation(builder, @"\f(1,2)"); // Fraction 1/2

        builder.Writeln("Equation 2:");
        InsertOfficeMathEquation(builder, @"\r(3,x)"); // Cube root of x

        builder.Writeln("Equation 3:");
        InsertOfficeMathEquation(builder, @"\i \su(n=1,5,n)"); // Integral with summation

        // Save the sample document.
        doc.Save(docPath);

        // Load the document (simulating an existing file).
        Document loadedDoc = new Document(docPath);

        // Extract all OfficeMath nodes.
        NodeCollection officeMathNodes = loadedDoc.GetChildNodes(NodeType.OfficeMath, true);
        List<string> equations = new List<string>();

        foreach (OfficeMath om in officeMathNodes)
        {
            // GetText provides a readable representation of the equation.
            string text = om.GetText().Trim();
            if (!string.IsNullOrEmpty(text))
                equations.Add(text);
        }

        // Write the extracted equations to a text file.
        string outputPath = Path.Combine(dataDir, "Equations.txt");
        File.WriteAllLines(outputPath, equations);

        // Validate that the output file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("Failed to create the equations report file.");

        // Optional: indicate completion.
        Console.WriteLine($"Extracted {equations.Count} equations to '{outputPath}'.");
    }

    // Inserts an EQ field, converts it to a real OfficeMath node, and removes the original field.
    private static void InsertOfficeMathEquation(DocumentBuilder builder, string args)
    {
        // Insert an EQ field.
        FieldEQ field = (FieldEQ)builder.InsertField(FieldType.FieldEquation, true);

        // Write the EQ arguments.
        builder.MoveTo(field.Separator);
        builder.Write(args);

        // Convert the field to OfficeMath.
        OfficeMath officeMath = field.AsOfficeMath();

        if (officeMath != null)
        {
            // Insert the OfficeMath node before the field start.
            field.Start.ParentNode.InsertBefore(officeMath, field.Start);
            // Remove the original field.
            field.Remove();

            // Move the builder after the inserted OfficeMath node.
            builder.MoveTo(officeMath);
        }
        else
        {
            // If conversion failed, keep the field as is.
            builder.MoveTo(field.Start);
        }

        // Add a paragraph break after the equation.
        builder.Writeln();
    }
}
