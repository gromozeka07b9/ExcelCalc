using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelCalc
{
    class Program
    {
        static void Main(string[] args)
        {
            if (args.Count() == 6)
            {
                string pathToSource = args[0];
                string pageName = args[1];
                ColumnsParameters columnsParameters = new ColumnsParameters();
                columnsParameters.columnNameBasicStart = args[2];
                columnsParameters.columnNameBasicEnd = args[3];
                columnsParameters.columnNameSpecialStart = args[4];
                columnsParameters.columnNameSpecialEnd = args[5];

                DataSet ds = FillTableFromFile(pathToSource, pageName);
                List<Employee> empl = FillEmployeeList(ds, pageName, columnsParameters);
                foreach (var item in empl)
                {
                    item.SocialPoints = item.SocialPoints / item.CountOfLead;
                    item.TechPoints = item.TechPoints / item.CountOfLead;
                }
                string pathToResult = Path.GetFileNameWithoutExtension(pathToSource) + "_" + DateTime.Now.ToShortDateString() + ".txt";
                exportToText(empl, pathToResult);
                Console.WriteLine("Успешно обработан!");
                Console.ReadLine();
            }
            else
            {
                Console.WriteLine("Некорректное количество параметров, должен быть формат: <путь к исходному файлу> <название листа> <имя колонки базовых начало> <имя колонки базовых конец> <имя колонки спец начало> <имя колонки спец конец>");
                Console.ReadLine();
            }
        }

        private static void exportToText(List<Employee> employeeList, string pathToTextFile)
        {
            StringBuilder sb = new StringBuilder();
            using (StreamWriter writer = new StreamWriter(pathToTextFile))
            {
                foreach (var employeeItem in employeeList)
                {
                    writer.WriteLine("{0}, базовые навыки:{1}, специальные навыки:{2}, количество оценивающих:{3}", employeeItem.Name, employeeItem.SocialPoints, employeeItem.TechPoints, employeeItem.CountOfLead);
                }
            }
        }

        private static List<Employee> FillEmployeeList(DataSet ds, string pageName, ColumnsParameters columnsParameters)
        {

            List<Employee> emplList = new List<Employee>();
            List<string> basicColumns = getColumnNames(ds.Tables[pageName].Columns, ds.Tables[pageName].Columns[columnsParameters.columnNameBasicStart].Ordinal, ds.Tables[pageName].Columns[columnsParameters.columnNameBasicEnd].Ordinal);
            List<string> specialColumns = getColumnNames(ds.Tables[pageName].Columns, ds.Tables[pageName].Columns[columnsParameters.columnNameSpecialStart].Ordinal, ds.Tables[pageName].Columns[columnsParameters.columnNameSpecialEnd].Ordinal);

            foreach (DataRow row in ds.Tables[pageName].Rows)
            {
                Console.WriteLine("basic points:");
                int basicPoints = summarizePoints(row, basicColumns);
                Console.WriteLine("special points:");
                int specialPoints = summarizePoints(row, specialColumns);
                string currentEmployeeName = row["Оцениваемый сотрудник"].ToString();
                Employee currentEmployee;
                if (!emplList.Exists(e => e.Name == currentEmployeeName))
                {
                    currentEmployee = new Employee();
                    currentEmployee.CountOfLead = 1;
                    currentEmployee.Name = currentEmployeeName;
                    emplList.Add(currentEmployee);
                }
                else
                {
                    currentEmployee = emplList.First(e => e.Name == currentEmployeeName);
                    currentEmployee.CountOfLead++;
                }
                currentEmployee.SocialPoints += basicPoints;
                currentEmployee.TechPoints += specialPoints;
            }
            return emplList;
        }

        private static int summarizePoints(DataRow row, List<string> columns)
        {
            int points = 0;
            foreach (var column in columns)
            {
                int count = getPointFromString(row[column].ToString());
                points += count;
                Console.WriteLine(row[column].ToString() + " = " + count);
            }
            return points;
        }

        private static int getPointFromString(string cellValue)
        {
            int result = 0;
            if(cellValue.Length>3)
            {
                string pointsDraft = cellValue.Substring(cellValue.Length - 4);
                string points = pointsDraft.Replace("(", "").Replace(")", "").Trim();
                result = Convert.ToInt32(points);
            }
            return result;
        }

        private static List<string> getColumnNames(DataColumnCollection columns, int columnFrom, int columnTo)
        {
            List<string> resultColumns = new List<string>();
            for (int i = columnFrom; i <= columnTo; i++)
            {
                resultColumns.Add(columns[i].ColumnName);
            }
            return resultColumns;                
        }

        private static DataSet FillTableFromFile(string path, string tableName)
        {
            DataSet ds = new DataSet();
            Dictionary<string, string> props = new Dictionary<string, string>();
            props["Provider"] = "Microsoft.Jet.OLEDB.4.0";
            props["Data Source"] = path;
            props["Extended Properties"] = "Excel 8.0";

            StringBuilder sb = new StringBuilder();
            foreach (KeyValuePair<string, string> prop in props)
            {
                sb.Append(prop.Key);
                sb.Append('=');
                sb.Append(prop.Value);
                sb.Append(';');
            }
            string properties = sb.ToString();

            using (OleDbConnection conn = new OleDbConnection(properties))
            {
                conn.Open();
                using (OleDbDataAdapter da = new OleDbDataAdapter(
                    "SELECT * FROM [" + tableName + "$]", conn))
                {
                    DataTable dt = new DataTable(tableName);
                    da.Fill(dt);
                    ds.Tables.Add(dt);
                }
            }
            return ds;
        }

        struct ColumnsParameters
        {
            public string columnNameBasicStart { get; internal set; }
            public string columnNameBasicEnd { get; internal set; }
            public string columnNameSpecialStart { get; internal set; }
            public string columnNameSpecialEnd { get; internal set; }
        }
    }
}
