using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace WindowsFormsApp1
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            var activities = new List<Activity>()
            {
                new Activity(){Name="Working", MondayHours=8, TuesdayHours=8, WednesdayHours=8},
                new Activity(){Name="Excercising", MondayHours=1, TuesdayHours=1, WednesdayHours=1},
                new Activity(){Name="Sleeping", MondayHours=6, TuesdayHours=7, WednesdayHours=8}
            };
            CreateSpreadsheet(activities);
            MessageBox.Show("All done");
        }

        private void CreateSpreadsheet(List<Activity> activities)
        {
            string spreadsheetPath = "activities.xlsx";
            File.Delete(spreadsheetPath);
            FileInfo spreadsheetInfo = new FileInfo(spreadsheetPath);

            ExcelPackage pck = new ExcelPackage();
            var activitiesWorksheet = pck.Workbook.Worksheets.Add("Activities");
            activitiesWorksheet.Cells["A1"].Value = "Name";
            activitiesWorksheet.Cells["B1"].Value = "Monday";
            activitiesWorksheet.Cells["C1"].Value = "Tuesday";
            activitiesWorksheet.Cells["D1"].Value = "Wendesday";
            
            activitiesWorksheet.Cells["A1:D1"].Style.Font.Bold = true;

            // populate spreadsheet with data
            int currentRow = 2;
            foreach(var activity in activities)
            {
                activitiesWorksheet.Cells["A" + currentRow.ToString()].Value = activity.Name;
                activitiesWorksheet.Cells["B" + currentRow.ToString()].Value = activity.MondayHours;
                activitiesWorksheet.Cells["C" + currentRow.ToString()].Value = activity.TuesdayHours;
                activitiesWorksheet.Cells["D" + currentRow.ToString()].Value = activity.WednesdayHours;

                currentRow++;
            }

            activitiesWorksheet.View.FreezePanes(2, 1);

            activitiesWorksheet.Cells["B" + (currentRow).ToString()].Formula = "SUM(B2:B" + (currentRow - 1).ToString() + ")";
            activitiesWorksheet.Cells["C" + (currentRow).ToString()].Formula = "SUM(C2:C" + (currentRow - 1).ToString() + ")";
            activitiesWorksheet.Cells["D" + (currentRow).ToString()].Formula = "SUM(D2:D" + (currentRow - 1).ToString() + ")";
            activitiesWorksheet.Cells["B" + (currentRow).ToString()].Style.Font.Bold = true;
            activitiesWorksheet.Cells["C" + (currentRow).ToString()].Style.Font.Bold = true;
            activitiesWorksheet.Cells["D" + (currentRow).ToString()].Style.Font.Bold = true;
            activitiesWorksheet.Cells["B" + (currentRow).ToString()].Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
            activitiesWorksheet.Cells["C" + (currentRow).ToString()].Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
            activitiesWorksheet.Cells["D" + (currentRow).ToString()].Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
            pck.SaveAs(spreadsheetInfo);
        } 
    }
}
