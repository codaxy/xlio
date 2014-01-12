XLIO
----
Xlio is a small C# library designed for reading and writing Excel files (.xlsx only).
It provides an elegant class system which feels natural for the developers.
It's very fast and has no external dependencies.

Check the following example to see how typical Xlio code looks like.

	var workbook = Workbook.ReadFile(@"C:\Work\Test.xlsx");
	var sheet = workbook.Sheets[0];
	sheet["A1"].Value = "Hello";
	sheet["A1"].Style.Fill = new CellFill
	{
		Foreground = new Color(255, 255, 0, 0),
		Pattern = FillPattern.Solid
	};
	sheet[0, 1].Value = "World";
	sheet["A3"].Value = "This is a merged cell with border.";
	sheet["A3"].Style.Alignment.WrapText = true;
	sheet["A3", "B5"].Merge();
	sheet["A3", "B5"].SetOutsideBorder(new BorderEdge
	{
		Style = BorderStyle.Thick,
		Color = Color.Black
	});

	sheet["A6"].Value = "Formatted:";
	var b6 = sheet["B6"];
	b6.Value = 5;
	b6.Style.Format = "#,#.00";

	for (var c = 0; c < 5; c++)
	{
		var bgColor = new Color((byte)(100 + c * 30), 255, (byte)(100 + c * 30));
		for (var r = 0; r < 50; r++)
			sheet[r, c].Style.Fill = CellFill.Solid(bgColor);
	}

	workbook.Save(@"C:\Work\Xlio.xlsx", XlsxFileWriterOptions.AutoFit);
	
More information available at the official [Xlio page](http://dox.codaxy.com/xlio/).