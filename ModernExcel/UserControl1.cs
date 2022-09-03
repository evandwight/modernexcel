using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using System.Diagnostics;

namespace ModernExcel
{
    public partial class MyUserControl : UserControl
    {
        public List<(MyAddress, string)> proposedNames;
        public int currentNameIndex = 0;
        public MyUserControl()
        {
            InitializeComponent();
        }
        public void setMessage(string text, bool isError = false)
        {
            this.messageBox.BackColor = Color.White;
            this.messageBox.ForeColor = isError ? Color.Red : Color.Black;
            this.messageBox.Text = text;
            this.messageBox.Visible = text.Length != 0;
        }
        public void goToNextName(int index)
        {
            if (index > this.proposedNames.Count - 1)
            {
                this.groupBox1.Visible = false;
                this.setMessage("Finished. No more names to update.");
                this.indexText.Text = "";
            }
            else
            {
                this.setMessage("");
                this.currentNameIndex = index;
                var currentName = this.proposedNames[index];
                var address = currentName.Item1;
                this.cellTextBox.Text = address.ToString();
                this.nameTextBox.Text = currentName.Item2;
                this.indexText.Text = String.Format("{0} of {1}", index + 1, this.proposedNames.Count);
                this.groupBox1.Visible = true;

                Excel.Workbook workbook = Globals.ThisAddIn.Application.ActiveWorkbook;
                var worksheet = workbook.Sheets[address.Sheet];
                worksheet.Select();
                worksheet.Cells[address.Row, address.Column].Select();
            }
        }



        private void applyButton_Click(object sender, EventArgs e)
        {
            Excel.Workbook workbook = Globals.ThisAddIn.Application.ActiveWorkbook;
            var currentName = this.proposedNames[this.currentNameIndex];
            var address = currentName.Item1;
            var newName = this.nameTextBox.Text;
            var worksheet = workbook.Sheets[address.Sheet];

            Regex rg = new Regex(@"^[_A-Za-z\d]+$");
            if (!rg.IsMatch(newName))
            {
                this.setMessage("Error: Name can only contain letters, numbers and underscore", true);
                return;
            }

            var registeredNames = ModernExcelRunner.getRegisteredNames(workbook);

            if (registeredNames.Values.Contains(newName))
            {
                this.setMessage("Error: Name already exists. Choose another name.", true);
                return;
            }

            workbook.Names.Add(newName, "=" + address.ToString());

            foreach (Excel.Worksheet sheet in workbook.Worksheets)
            {
                // UseColumnRowNames=false because we don't want to use whole column/row names to name intersections
                // AppendLast=false because we only want to replace names in newNames, not that any will be more recently defined
                sheet.Cells.ApplyNames(new string[] { newName }, true, false, false, false, Excel.XlApplyNamesOrder.xlRowThenColumn, false);
            }

            this.goToNextName(this.currentNameIndex + 1);
        }

        private void runButton_Click(object sender, EventArgs e)
        {
            this.setMessage("Analyzing, this may take a while.");
            this.proposedNames = ModernExcelRunner.getProposedNames().ToList();
            if (this.proposedNames.Count == 0)
            {
                this.setMessage("No cells to name found");
            }
            else
            {
                this.goToNextName(0);
            }
        }

        private void skipButton_Click(object sender, EventArgs e)
        {
            this.goToNextName(this.currentNameIndex + 1);
        }
    }
}
