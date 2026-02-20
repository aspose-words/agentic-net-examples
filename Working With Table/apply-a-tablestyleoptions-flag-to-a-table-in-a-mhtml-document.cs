// Load an existing MHTML document
string inputPath = @"C:\Docs\input.mhtml";
Aspose.Words.Document doc = new Aspose.Words.Document(inputPath);

// Get the first table in the document (adjust index as needed)
Aspose.Words.Tables.Table table = doc.FirstSection.Body.Tables[0];

// Optionally set a built‑in table style
table.StyleIdentifier = Aspose.Words.StyleIdentifier.TableGrid;

// Apply desired style options (e.g., first row and row banding)
table.StyleOptions = Aspose.Words.Tables.TableStyleOptions.FirstRow |
                     Aspose.Words.Tables.TableStyleOptions.RowBands;

// Save the document back to MHTML format
Aspose.Words.Saving.HtmlSaveOptions saveOptions = new Aspose.Words.Saving.HtmlSaveOptions(Aspose.Words.SaveFormat.Mhtml);
doc.Save(@"C:\Docs\output.mhtml", saveOptions);
