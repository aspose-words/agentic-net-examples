using System;
using System.Drawing;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;

// Load the existing DOC/DOCX document (lifecycle rule: load)
Document doc = new Document("InputDocument.docx");

// Define the style name we want to target
string targetStyleName = "MyStyle";

// Select all Run nodes in the document, filter by the specified style using LINQ
var runsWithStyle = doc.GetChildNodes(NodeType.Run, true)
                       .Cast<Run>()
                       .Where(r => string.Equals(r.Font.StyleName, targetStyleName, StringComparison.Ordinal));

// Change the font color of each matching run
foreach (Run run in runsWithStyle)
{
    run.Font.Color = Color.Red; // Set desired color
}

// Save the modified document (lifecycle rule: save)
doc.Save("OutputDocument.docx");
