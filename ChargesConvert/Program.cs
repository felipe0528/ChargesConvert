using System;
using System.Data;
using System.IO;
using Newtonsoft.Json;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;

class Program
{
    static void Main(string[] args)
    {
        // Read the data from the Excel XLS file
        DataTable data = ReadExcelFile("C:\\Users\\felip\\OneDrive\\Trabajos\\varner.trip\\Projects\\ChargesConvert\\ChargesConvert\\bin\\Debug\\netcoreapp3.1\\Charges.xls");

        Charge[] charges = new Charge[data.Rows.Count];
        for (int i = 0; i < data.Rows.Count; i++)
        {
            DataRow row = data.Rows[i];
            charges[i] = new Charge
            {
                id = row["Value"] + ":" +Guid.NewGuid().ToString(),
                Type = new Type
                {
                    Value = Convert.ToInt32(row["Value"]),
                    Label = row["Type"].ToString().Trim()
                },
                Code = row["Code"].ToString().Trim(),
                TISCode = row["TIS Code"].ToString().Trim(),
                Description = row["Description"].ToString().Trim()
            };
        }

        // Serialize the array of Charge objects to a JSON string
        string json = JsonConvert.SerializeObject(charges, Formatting.Indented);

        // Write the JSON string to a file
        File.WriteAllText("charges.json", json);

        Console.WriteLine("Conversion from Excel XLS to JSON has been completed successfully.");
        Console.ReadLine();
    }

    static DataTable ReadExcelFile(string filePath)
    {
        DataTable data = new DataTable();
        data.Columns.Add("Code");
        data.Columns.Add("Description");
        data.Columns.Add("TIS Code");
        data.Columns.Add("Type");
        data.Columns.Add("Value");

        using (FileStream file = new FileStream(filePath, FileMode.Open, FileAccess.Read))
        {
            HSSFWorkbook workbook = new HSSFWorkbook(file);
            ISheet sheet = workbook.GetSheetAt(0);

            for (int row = 1; row <= sheet.LastRowNum; row++)
            {
                IRow excelRow = sheet.GetRow(row);
                if (excelRow == null) continue;

                DataRow dataRow = data.NewRow();
                dataRow["Code"] = excelRow.GetCell(0)?.ToString();
                dataRow["Description"] = excelRow.GetCell(1)?.ToString();
                dataRow["TIS Code"] = excelRow.GetCell(2)?.ToString();
                dataRow["Type"] = excelRow.GetCell(3)?.ToString();
                dataRow["Value"] = excelRow.GetCell(4)?.ToString();
                data.Rows.Add(dataRow);
            }
        }

        return data;
    }
}

class Charge
{
    public string id { get; set; }
    public Type Type { get; set; }
    public string Code { get; set; }
    public string TISCode { get; set; }
    public string Description { get; set; }
}

class Type
{
    public int Value { get; set; }
    public string Label { get; set; }
}
