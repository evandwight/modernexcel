using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using System.Text.RegularExpressions;

namespace ModernExcel
{
    public class MyAddress
    {
        public int Row;
        public int Column;
        public string Sheet;
        public MyAddress(Excel.Range cell) : this(cell.Worksheet.Name, cell.Row, cell.Column)
        {
        }
        public MyAddress(string sheet, int row, int column)
        {
            this.Row = row;
            this.Column = column;
            this.Sheet = sheet;
        }

        // thank you Cristian Lupascu
        // https://stackoverflow.com/questions/10373561/convert-a-number-to-a-letter-in-c-sharp-for-use-in-microsoft-excel
        public static string GetColumnName(int index)
        {
            index -= 1; // excel is 1 indexed

            if (index < 0)
            {
                throw new Exception("index must be greater than 1");
            }

            var value = "";

            for (int i = 0; i < 10; i++)
            {
                value = (char)('A' + index % 26) + value;
                if (index < 26)
                {
                    break;
                }
                index = index / 26 - 1;
            }

            return value;
        }
        public override string ToString()
        {
            return this.Sheet + "!$" + MyAddress.GetColumnName(this.Column) + "$" + this.Row.ToString();
        }
        public override bool Equals(object obj)
        {
            var other = obj as MyAddress;
            if (other is null)
            {
                return false;
            }
            return this.Column == other.Column
                && this.Row == other.Row
                && this.Sheet == other.Sheet;
        }
        public override int GetHashCode() => (this.Column, this.Row, this.Sheet).GetHashCode();
    }
    public class ModernExcelRunner
    {
        public static Dictionary<string, TValue> convertKeysToString<TKey, TValue>(Dictionary<TKey, TValue> d)
        {
            return d.ToArray().ToDictionary(keySelector: m => m.Key.ToString(), elementSelector: m => m.Value);

        }
        public static IEnumerable<(MyAddress, string)> getProposedNames()
        {
            Excel.Workbook workbook = Globals.ThisAddIn.Application.ActiveWorkbook;

            var precedents = getFormulaPrecedents(workbook);

            Debug.WriteLine("All formula references");
            precedents.ForEach(e => Debug.WriteLine(e.Item1 + " - " + String.Join(", ", e.Item2)));

            IEnumerable<(MyAddress, string)> interestingCells = precedents
                .SelectMany(e => e.Item2)
                .Distinct()
                .Select(e => (e, ModernExcelRunner.guessName(e, workbook)))
                .Where(e => !(e.Item2 is null));

            Debug.WriteLine("Interesting cells");
            interestingCells.ToList().ForEach(e => Debug.WriteLine(e.Item1 + " - " + e.Item2));

            Dictionary<MyAddress, string> registeredNames = ModernExcelRunner.getRegisteredNames(workbook);


            var unnamedCells = interestingCells.Where(e => !registeredNames.ContainsKey(e.Item1));

            Debug.WriteLine("Interesting unnamed cells");
            unnamedCells.ToList().ForEach(e => Debug.WriteLine(e.Item1 + " - " + e.Item2));

            return unnamedCells;
        }
        public static string guessName(MyAddress cell, Excel.Workbook workbook)
        {
            var label = cell.Column > 1 ? workbook.Sheets[cell.Sheet].Cells[cell.Row, cell.Column - 1] : null;
            if (label?.Value is null || label?.Value.ToString().Trim() == "")
            {
                return null;
            }
            else
            {
                return ModernExcelRunner.textToSnakeCase(label.Value.ToString());
            }
        }
        public static List<MyAddress> getCellPrecendents(Excel.Range cell)
        {
            var list = new List<MyAddress>();
            if (cell == null || !cell.HasFormula)
            {
                return list;
            }
            Excel.Areas precedents = null;
            try
            {
                precedents = cell.DirectPrecedents.Areas;
            }
            catch (Exception e)
            {
                Debug.WriteLine(e);
                return list;
            }

            foreach (Excel.Range precedent in precedents)
            {
                list.Add(new MyAddress(precedent));
            }

            return list;
        }

        public static Dictionary<MyAddress, string> getRegisteredNames(Excel.Workbook workbook)
        {
            Excel.Names named = workbook.Names;
            Dictionary<MyAddress, string> allNames = new Dictionary<MyAddress, string>();
            foreach (Excel.Name name in workbook.Names)
            {
                allNames[new MyAddress(name.RefersToRange)] = name.Name;
            }
            foreach (Excel.Worksheet worksheet in workbook.Worksheets)
            {
                foreach (Excel.Name name in worksheet.Names)
                {
                    allNames[new MyAddress(name.RefersToRange)] = name.Name;
                }
            }
            return allNames;
        }
        public static List<(MyAddress, IEnumerable<MyAddress>)> getFormulaPrecedents(Excel.Workbook workbook)
        {
            const int MAX_ROWS = 1048576;
            var allPrecendents = new List<(MyAddress, IEnumerable<MyAddress>)>();
            foreach (Excel.Worksheet sheet in workbook.Worksheets)
            {
                Excel.Range usedRange = sheet.UsedRange;
                foreach (Excel.Range cell in usedRange.Cells)
                {
                    if (cell.HasFormula)
                    {
                        var myCell = new MyAddress(cell);
                        IEnumerable<MyAddress> precedents = ModernExcelRunner.getCellPrecendents(cell);

                        // Remove precedents probably in a table
                        Excel.Range above = cell.Row > 1 ? sheet.Cells[cell.Row - 1, cell.Column] : null;
                        if (above?.FormulaR1C1 == cell.FormulaR1C1)
                        {
                            precedents = precedents.Intersect(ModernExcelRunner.getCellPrecendents(above));
                        }
                        Excel.Range below = cell.Row < MAX_ROWS ? sheet.Cells[cell.Row + 1, cell.Column] : null;
                        if (below?.FormulaR1C1 == cell.FormulaR1C1)
                        {
                            precedents = precedents.Intersect(ModernExcelRunner.getCellPrecendents(above));
                        }
                        allPrecendents.Add((myCell, precedents));
                    }
                }
            }
            return allPrecendents;
        }
        public static string textToSnakeCase(string text)
        {
            text = text.ToLower();
            Regex rg = new Regex(@"[a-z\d]+");

            return String.Join("_", (from Match m in rg.Matches(text) select m.Value).ToList());
        }
    }
}