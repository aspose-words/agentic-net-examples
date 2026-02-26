// Load an existing DOTX template (or create a new document if needed)
var doc = new Aspose.Words.Document("Template.dotx");

// Create a DocumentBuilder to work with the document
var builder = new Aspose.Words.DocumentBuilder(doc);

// Insert a horizontal rule shape at the current cursor position
Aspose.Words.Drawing.Shape hrShape = builder.InsertHorizontalRule();

// Optional: customize the horizontal rule appearance
hrShape.HorizontalRuleFormat.Alignment = Aspose.Words.Drawing.HorizontalRuleAlignment.Center;
hrShape.HorizontalRuleFormat.WidthPercent = 70;   // 70% of the page width
hrShape.HorizontalRuleFormat.Height = 3;          // 3 points high
hrShape.HorizontalRuleFormat.Color = System.Drawing.Color.Blue;
hrShape.HorizontalRuleFormat.NoShade = true;     // Solid color, no 3D shading

// Save the modified document back to DOTX format
doc.Save("Result.dotx");
