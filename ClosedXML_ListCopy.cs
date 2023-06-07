using ClosedXML.Excel;

public class Item
{
    public string TsNumber { get; set; } = "";
    public int Quantity { get; set; }
    public string LocNumber { get; set; } = "";
}

public class ExcelOperator
{
    public void OperateExcel(string filePath)
    {
        var workbook = new XLWorkbook(filePath);

        // Sheet1からデータを読み込む
        var sheet1 = workbook.Worksheet("Sheet1");
        List<Item> items = new List<Item>();
        for (int i = 1; i <= 2; i++)
        {
            var row = sheet1.Row(i);
            items.Add(new Item()
            {
                TsNumber = row.Cell(1).GetValue<string>(),
                Quantity = row.Cell(2).GetValue<int>(),
                LocNumber = row.Cell(3).GetValue<string>()
            });
        }

        // List配列をTS図番の文字列で昇順に並べ替える
        items = items.OrderBy(item => item.TsNumber).ToList();

        // Sheet2が存在しない場合、新規に作成する
        IXLWorksheet sheet2;
        if (!workbook.TryGetWorksheet("Sheet2", out sheet2))
        {
            sheet2 = workbook.AddWorksheet("Sheet2");
        }

        // List配列の内容をSheet2に転記する
        for (int i = 0; i < items.Count; i++)
        {
            var item = items[i];
            var row = sheet2.Row(i + 1);
            row.Cell(1).Value = item.Quantity;
            row.Cell(2).Value = item.TsNumber;
            row.Cell(3).Value = item.LocNumber;
        }

        // 変更を保存する
        workbook.Save();
    }
}

class Program
{
    static void Main()
    {
        ExcelOperator excelOperator = new ExcelOperator();
        excelOperator.OperateExcel("BaseExcel.xlsx");
    }
}
