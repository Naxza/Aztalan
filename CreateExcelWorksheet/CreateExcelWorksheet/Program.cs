using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;


using System.Reflection;
using Microsoft.Office.Interop.Excel;
using System.Text.RegularExpressions;
using System.IO;
using System.Runtime.InteropServices;

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
        Console.WriteLine("Please enter file name: ");
        //string file = Console.ReadLine();
        string fileName = Console.ReadLine(); //user enters file name
        string worksheetName = "Shipment - ByCell per week";
        

        Application xlApp = new Application();//creates a new reference to excel
        if (xlApp == null)
        {
            Console.WriteLine("EXCEL could not be started. Check that your office installation and project references are correct.");
            return;
        }
        
        Workbook wb=null;//creates a wb reference
        while (wb == null)
        {
            if (!File.Exists(loc + fileName + ".xlsm"))
            {            
                Console.Write("File doesn't exist.  please try again\n");
                Console.WriteLine("Please enter file name: ");
                fileName = Console.ReadLine();
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
        }
        xlApp.Visible = true;
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
        
        ws = wb.Sheets[2];//sets ws to the second sheet as the wb.sheets[worksheetName] code was not working at the time
        xlApp.ActiveWindow.Zoom = 50;//sets the zoom of the page out so you can see the chart as it generates
        xlApp.DisplayAlerts = true;//allows excel to show any errors or alerts
        if (ws == null)
        {
            Console.WriteLine("\nWorksheet could not be created. Check that your office installation and project references are correct.");
        }
        chart(ws, "C1", "Z2", chartName);

        bool deleted = false;

        /*AFTER CREATING AND SAVING THE CHART THE CHART IS DELETED
        *The reasoning behind this decision is so that if charts are generated constantly the file size will
        *continue to grow and take up more space than needed, this way the chart is saved and stored as an immage.
        *the immages can then be pulled to show a history of the past week.
        */
        try
        {
            ChartObject chartPageDelete = ws.ChartObjects("Chart 1");
            chartPageDelete.Delete();
            deleted = true;
        }
        catch
        {
            Console.WriteLine("\nChart with this name could not be found");
            //throw new Exception("Chart with this name could not be found");
        }
        finally
        {
            Console.WriteLine("\nthe chart was " + (deleted ? "deleted" : "not deleted"));
        }


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
            wb.Close();//closes the workbook
            
        }
        Marshal.ReleaseComObject(xlApp);//releases all Com objects
        xlApp.Quit();//closes this instance of excel

    }


    //function to create charts, can be called for multiple charts
    static void chart(Worksheet thisWorkSheet,string x, string y, string name)
    {
        ChartObjects newCharts = (ChartObjects)thisWorkSheet.ChartObjects(Type.Missing);//creates a chart object
        ChartObject myChart = (ChartObject)newCharts.Add(10, 80, 1350, 550);//position of chart(pos X, pos y, width, height)
        Chart chartPage = myChart.Chart;
        Range chartRange = thisWorkSheet.get_Range(x,y);
        chartPage.Rotation = 0;
        chartPage.ChartType = XlChartType.xl3DColumnClustered;
        Axis xAxis = (Axis)chartPage.Axes(XlAxisType.xlCategory, XlAxisGroup.xlPrimary);
        xAxis.TickLabelPosition =XlTickLabelPosition.xlTickLabelPositionNone;
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
            //simple check statement to add more colums.
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
        
        chartPage.HasTitle = true;//makes the chart have a title
        chartPage.ChartTitle.Text = "Sales Data, Week of "+thisWorkSheet.get_Range("a2","a2").Value;//title of the chart
        chartPage.ChartTitle.Font.Size = 45;//sets the font size of the title
        chartPage.ChartTitle.Font.Color = XlRgbColor.rgbGoldenrod;//color of title
        chartPage.ChartStyle = 42;//style of chart
        DateTime week = DateTime.Today;
        string holder = Format(week.ToString("d"));//gets date to add to the immage that is being saved
        //EVERYTHING THAT HAS AN EFFECT ON THE CHART SHOULD HAPPEN BEFORE THIS!
        chartPage.Export(loc + name +holder+ ".jpg", "JPG", false);//takes the newly generated chart and saves it as a jpg
        //THIS IS BASICALLY THE RETURN STATEMENT  /\  DO NOT DELETE!
        //      (not really but it save the immage)
        
    }


    //formats date so that it can be added to a file name
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

    //saves changes to the file through a function so it may be called multiple times
    static void save(Application Excel, Workbook THING, string fileNomen)
    {
        string path = loc;
        Excel.DisplayAlerts = false;
        THING.Application.DefaultFilePath = path;
        THING.SaveAs(fileNomen, Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookDefault, Type.Missing, Type.Missing, true, false, XlSaveAsAccessMode.xlNoChange, XlSaveConflictResolution.xlLocalSessionChanges, Type.Missing, Type.Missing);
        THING.Close();
    }
}