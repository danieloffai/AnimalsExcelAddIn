using Microsoft.Office.Tools.Ribbon;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Excel;
using Microsoft.Office.Interop.Excel;

namespace ExcelAddInZoo
{
    public partial class Ribbon
    {
        private void Ribbon_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void btnStart_Click(object sender, RibbonControlEventArgs e)
        {
            int animalNumber;

            using (MessageBoxAnimalNumber messageBoxAnimalNumber = new MessageBoxAnimalNumber())
            {
                messageBoxAnimalNumber.BringToFront();
                messageBoxAnimalNumber.ShowDialog();

                animalNumber = messageBoxAnimalNumber.AnimalNumber;
            }

            PopulateAnimals(animalNumber);
        }

        private void PopulateAnimals(int animalNumber)
        {
            var animals = APIsession.GetAnimals(animalNumber);

            Excel.Workbook currentWB = Globals.ThisAddIn.Application.ActiveWorkbook;
            Excel.Worksheet activeWorksheet = Globals.ThisAddIn.Application.ActiveSheet;
            Excel.Range startingCell = Globals.ThisAddIn.Application.ActiveCell;

            activeWorksheet.Rows.Clear();

            int activeCellRow = startingCell.Row;
            int activeCellColumn = startingCell.Column;

            PopulateAnimalAttributes(activeWorksheet, startingCell, activeCellRow);

            PopulateAnimalInfo(animals, activeWorksheet, startingCell, ref activeCellRow, ref activeCellColumn);

            activeWorksheet.Columns.AutoFit();

            PopulateChart(currentWB, activeWorksheet, startingCell, activeCellRow, activeCellColumn);
        }

        private void PopulateChart(Excel.Workbook currentWB, Excel.Worksheet activeWorksheet, Range startingCell, int activeCellRow, int activeCellColumn)
        {
            string lifespanStart = RangeAddress(activeWorksheet.Cells[startingCell.Row + 8, startingCell.Column + 1]);
            string lifespanEnd = RangeAddress(activeWorksheet.Cells[activeCellRow + 8, activeCellColumn - 1]);

            string namesStart = RangeAddress(activeWorksheet.Cells[startingCell.Row, startingCell.Column + 1]);
            string namesEnd = RangeAddress(activeWorksheet.Cells[activeCellRow, activeCellColumn - 1]);

            Excel.Range lifespan = activeWorksheet.Range[$"{lifespanStart}:{lifespanEnd}"];
            Excel.Range names = activeWorksheet.Range[$"{namesStart}:{namesEnd}"];

            Excel.Chart chart = (Excel.Chart)currentWB.Charts.Add(Type.Missing, Type.Missing, 1, Type.Missing);
            chart.ChartTitle.Text = "Average lifespan";

            chart.ChartType = Excel.XlChartType.xlPie;

            var seriesX = (Series)chart.SeriesCollection(1);
            seriesX.Values = lifespan;
            seriesX.XValues = names;
        }

        private static void PopulateAnimalAttributes(Excel.Worksheet activeWorksheet, Range activeCell, int activeCellRow)
        {
            List<string> animalAttributes = new List<string>() { "Name", "Latin Name", "Animal Type", "Active Time", "Min Length", "Max Length",
                "Min Weight", "Max Weight", "Lifespan", "Habitat", "Diet", "Geo Range"};

            foreach (string attribute in animalAttributes)
            {
                activeWorksheet.Cells[activeCellRow, activeCell.Column].Value = attribute;
                activeWorksheet.Cells[activeCellRow, activeCell.Column].Font.Bold = true;
                activeCellRow++;
            }
        }

        private void PopulateAnimalInfo(List<ZooAnimal> animals, Excel.Worksheet activeWorksheet, Range startingCell, ref int activeCellRow, ref int activeCellColumn)
        {
            for (int i = 0; i < animals.Count + 1; i++)
            {
                activeCellColumn++;
                activeCellRow = startingCell.Row;

                if (animals.Count > i)
                {
                    for (int j = 0; j < 1; j++)
                    {
                        var animal = animals[i];
                        activeWorksheet.Cells[activeCellRow++, activeCellColumn].Value = animal.name;
                        activeWorksheet.Cells[activeCellRow++, activeCellColumn].Value = animal.latin_name;
                        activeWorksheet.Cells[activeCellRow++, activeCellColumn].Value = animal.animal_type;
                        activeWorksheet.Cells[activeCellRow++, activeCellColumn].Value = animal.active_time;
                        activeWorksheet.Cells[activeCellRow++, activeCellColumn].Value = animal.length_min;
                        activeWorksheet.Cells[activeCellRow++, activeCellColumn].Value = animal.length_max;
                        activeWorksheet.Cells[activeCellRow++, activeCellColumn].Value = animal.weight_min;
                        activeWorksheet.Cells[activeCellRow++, activeCellColumn].Value = animal.weight_max;
                        activeWorksheet.Cells[activeCellRow++, activeCellColumn].Value = animal.lifespan;
                        activeWorksheet.Cells[activeCellRow++, activeCellColumn].Value = animal.habitat;
                        activeWorksheet.Cells[activeCellRow++, activeCellColumn].Value = animal.diet;
                        activeWorksheet.Cells[activeCellRow++, activeCellColumn].Value = animal.geo_range;
                    }
                }
                else
                {
                    activeWorksheet.Cells[activeCellRow, activeCellColumn].Value = "AVERAGE";

                    string lengthMinAveStart = RangeAddress(activeWorksheet.Cells[startingCell.Row + 4, startingCell.Column + 1]);
                    string lengthMinAveEnd = RangeAddress(activeWorksheet.Cells[activeCellRow + 4, activeCellColumn - 1]);
                    activeWorksheet.Cells[activeCellRow + 4, activeCellColumn].Value = $"=AVERAGE({lengthMinAveStart}:{lengthMinAveEnd})";

                    string lengthMaxAveStart = RangeAddress(activeWorksheet.Cells[startingCell.Row + 5, startingCell.Column + 1]);
                    string lengthMaxAveEnd = RangeAddress(activeWorksheet.Cells[activeCellRow + 5, activeCellColumn - 1]);
                    activeWorksheet.Cells[activeCellRow + 5, activeCellColumn].Value = $"=AVERAGE({lengthMaxAveStart}:{lengthMaxAveEnd})";

                    string weightMinAveStart = RangeAddress(activeWorksheet.Cells[startingCell.Row + 6, startingCell.Column + 1]);
                    string weightMinAveEnd = RangeAddress(activeWorksheet.Cells[activeCellRow + 6, activeCellColumn - 1]);
                    activeWorksheet.Cells[activeCellRow + 6, activeCellColumn].Value = $"=AVERAGE({weightMinAveStart}:{weightMinAveEnd})";

                    string weightMaxAveStart = RangeAddress(activeWorksheet.Cells[startingCell.Row + 7, startingCell.Column + 1]);
                    string weightMaxAveEnd = RangeAddress(activeWorksheet.Cells[activeCellRow + 7, activeCellColumn - 1]);
                    activeWorksheet.Cells[activeCellRow + 7, activeCellColumn].Value = $"=AVERAGE({weightMaxAveStart}:{weightMaxAveEnd})";

                    string lifespanStart = RangeAddress(activeWorksheet.Cells[startingCell.Row + 8, startingCell.Column + 1]);
                    string lifespanEnd = RangeAddress(activeWorksheet.Cells[activeCellRow + 8, activeCellColumn - 1]);
                    activeWorksheet.Cells[activeCellRow + 8, activeCellColumn].Value = $"=AVERAGE({lifespanStart}:{lifespanEnd})";
                }
            }
        }

        public string RangeAddress(Excel.Range rng)
        {
            return rng.get_AddressLocal(false, false, Excel.XlReferenceStyle.xlA1);
        }
    }
}
