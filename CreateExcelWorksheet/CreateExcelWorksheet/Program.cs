using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;


using System.Reflection;
using Microsoft.Office.Interop.Excel;
using System.Text.RegularExpressions;
using System.IO;

public class CreateExcelWorksheet
{
    static string loc = Directory.GetParent(Directory.GetCurrentDirectory()).Parent.FullName + "/";
    
    static void Main()
    {
        Console.WriteLine(loc);
        Console.Write("\nchecking authorization level");//quickly check what the user's level is
        /*
        CHANGE THIS TO A LOGIN  \/
        */
        string Authorization = "admin";//checks login 
        /*
        CHANGE THIS TO A LOG IN /\
        */

        Console.Write("begining check process");//
        string chartName = "chart_";        
        Console.Write("Attempting to open file\n");//"please enter a name for the file:");
        string fileName = "SalesByWeekly";//Console.ReadLine();
        string worksheetName = "Shipment - ByCell per week";
        

        Application xlApp = new Application();//creates a new reference to excel
        if (xlApp == null)
        {
            Console.WriteLine("EXCEL could not be started. Check that your office installation and project references are correct.");
            return;
        }
        xlApp.Visible = true;
        Workbook wb;
        if (!File.Exists(loc + fileName + ".xlsm"))
        {

            if (Authorization.CompareTo("admin") == 0)
            {
                Console.Write("File doesn't exist.  creating new");
                
                wb = xlApp.Workbooks.Add(XlWBATemplate.xlWBATWorksheet);
            }
            wb = null;

        }
        else
        {
            Console.Write("File exists!  opening file");
            if (Authorization.CompareTo("admin") != 0)
            {
                xlApp.DisplayAlerts = false;
            }
            else
            {
            }
            wb = xlApp.Workbooks.Open(loc + fileName + ".xlsm", Type.Missing, false);

        }
        Boolean found = false;
        Worksheet ws=null;
        //xlApp.DisplayAlerts = false;
        if (wb != null)
        {
            foreach (Worksheet sheet in wb.Sheets)
            {
                // Check the name of the current sheet
                if (sheet.Name == "Shipment - ByCell per week")
                {
                    found = true;
                    Console.Write("worksheet was found\n");
                    break; // Exit the loop now
                }
            }
            //ws = (Worksheet)wb.Worksheets[worksheetName];
        }

        else
        {
            ws = null;
        }
        if (found)
        {
            // Reference it by name
            ws = wb.Sheets[worksheetName];
            chart(ws, "C1", "N2",chartName);
        }
        ws = wb.Sheets[2];
        xlApp.ActiveWindow.Zoom = 50;
        xlApp.DisplayAlerts = true;
        if (ws == null)
        {
            Console.WriteLine("\nWorksheet could not be created. Check that your office installation and project references are correct.");
        }
        chart(ws, "C1", "Z2", chartName);
        //ws.Shapes.Item(mainChart).Top = 500;//<-cant interact with objects once excel is open causes errors :(
        //begining manipulation/transfer attempt
        xlApp.DisplayAlerts = false;
        

        //saving the file
        Console.WriteLine("would you like to save? (Y or N)");
        string userIn = Console.ReadLine();
        if (userIn.Equals("Y") || userIn.Equals("y"))
        {
            save(xlApp, wb, fileName);
        }
        else
        {
            
        }
    }

    static void chart(Worksheet thisWorkSheet,string x, string y, string name)
    {
        ChartObjects newCharts = (ChartObjects)thisWorkSheet.ChartObjects(Type.Missing);
        ChartObject myChart = (ChartObject)newCharts.Add(10, 80, 1350, 550);//(pos X, pos y, width, height)
        Chart chartPage = myChart.Chart;
        Range chartRange = thisWorkSheet.get_Range(x,y);
        chartPage.Rotation = 0;
        chartPage.ChartType = XlChartType.xl3DColumnClustered;
        Axis xAxis = (Axis)chartPage.Axes(XlAxisType.xlCategory, XlAxisGroup.xlPrimary);
        xAxis.TickLabelPosition =XlTickLabelPosition.xlTickLabelPositionNone;
        xAxis.HasTitle = true;
        xAxis.AxisTitle.Text = " ";//NEEDS TO BE AXIS NOT AXIS TITLE SOMEHOW
        int startCol = 3;
        int startRow = 1;
        int maxRow = 2;
        int maxCol = 14; //With this value the loop below will iterate until column 9 (inclusive)
        for (int i = startCol; i <= maxCol; i++)
        {
            Range holder1 = thisWorkSheet.Cells[1, i];
            Range holder2 = thisWorkSheet.Cells[2, i];
            Range holder3 = thisWorkSheet.get_Range(holder1, holder2);
            bool check = false;
            for (int j = startRow; j <= maxRow; j++) {
                
                Range currentRange = (Range)thisWorkSheet.Cells[j, i];//iterates through the data in the excel document
                if (currentRange.Value2 != null)
                {
                    check = true;
                    string curVal = currentRange.Value2.ToString();
                    //Console.WriteLine("\n :loading value "+i+" part "+j);
                }
            }
            if (check)
            {
                //Series newSeries;
                SeriesCollection seriesCollect = chartPage.SeriesCollection();
                //chartPage.Axes().text += holder1;
                Series seriesI= seriesCollect.NewSeries();
                seriesI.XValues = holder1;
                seriesI.Values = holder2;
                seriesI.HasDataLabels = true;
                seriesI.HasLeaderLines = true;
                //Console.WriteLine (xAxis.TickLabels.Name);
                seriesI.Name = holder1.Value;

            }
        }
        //chartPage.SetSourceData(chartRange, Type.Missing);
        
        chartPage.HasTitle = true;
        chartPage.ChartTitle.Text = "Sales Data, Week of "+thisWorkSheet.get_Range("a2","a2").Value;//title of the chart
        chartPage.ChartTitle.Font.Size = 45;
        chartPage.ChartTitle.Font.Color = XlRgbColor.rgbGoldenrod;
        chartPage.ChartStyle = 42;
        DateTime week = DateTime.Today;
        string holder = Format(week.ToString("d"));
        //EVERYTHING THAT HAS AN EFFECT ON THE CHART SHOULD HAPPEN BEFORE THIS!
        //chartPage.Export(loc+name+(thisWorkSheet.get_Range("a2", "a2").Value).toString()+".jpg", "JPG", false);//takes the newly generated chart and saves it as a jpg
        chartPage.Export(loc + name +holder+ ".jpg", "JPG", false);//takes the newly generated chart and saves it as a jpg

        //THIS IS BASICALLY THE RETURN STATEMENT  /\  DO NOT DELETE!
    }

    static string Format(string holder)
    {
        char[] characters = holder.ToCharArray();
        for (int index = 0; index < characters.Length; index++)
        {
            if (characters[index] == '/')
            {
                characters[index] = '-';
            }
        }
        return new string(characters);
    }
    static void save(Application Excel, Workbook THING, string fileNomen)
    {
        string path = @"C:\Users\Sam Kromm\Documents\Aztalalalalalan-Project-Software-Engineering-\CreateExcelWorksheet\CreateExcelWorksheet";
        Excel.DisplayAlerts = false;
        THING.Application.DefaultFilePath = path;
        THING.SaveAs(fileNomen, Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookDefault, Type.Missing, Type.Missing, true, false, XlSaveAsAccessMode.xlNoChange, XlSaveConflictResolution.xlLocalSessionChanges, Type.Missing, Type.Missing);
        THING.Close();
    }
}

