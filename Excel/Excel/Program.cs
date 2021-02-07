using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;
using _Excel = Microsoft.Office.Interop.Excel;

namespace Excel {
    class Program {

        public List<string> singleNames = new List<string>();

        private string path = @"C:\Users\gabri\Desktop\csharp.xlsx";
        private Application app = new _Excel.Application();
        private Workbook book;
        private Worksheet sheet;

        public Program() {
            Initialize();
            GetNames();
            book.Close();
        }

        public void Initialize() {
            book = app.Workbooks.Open(path, Notify:false);
            sheet = book.Worksheets[1];
        }

        public void GetNames() {
            int? nameColumn = null;
            int columnX = 1;
            int columnY = 2;
            while (true) {
                if (sheet.Cells[1, columnX].Value2 == null) break;
                else if (sheet.Cells[1, columnX].Value2 == "Nomes") nameColumn = columnX;
                columnX++;
            }
            if (nameColumn == null) Console.WriteLine("I didn't find any column for names, please insert one");
            else {
                Console.WriteLine("Found the names column");
                while (true) {
                    if (sheet.Cells[columnY, nameColumn].Value2 == null) break;//found a cell with nothing, must be the endline, whatever
                    InsertSingleNames(sheet.Cells[columnY, nameColumn].Value2);
                    columnY++;
                }
            }
            Console.Write("{ ");
            for (int i = 0; i < singleNames.Count; i++) {
                Console.Write(singleNames[i] + ", ");
            }
            Console.Write(" }");
        }

        public void InsertSingleNames(string name) {
            for (int i = 0; i < singleNames.Count; i++) {
                if (singleNames[i] == name) return;
            }

            singleNames.Add(name);
        }

        static void Main(string[] args) {
            Program main = new Program();
            Console.ReadLine();
        }
    }
}
