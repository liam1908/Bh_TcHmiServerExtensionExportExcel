//-----------------------------------------------------------------------
// <copyright file="SE_Test.cs" company="Beckhoff Automation GmbH & Co. KG">
//     Copyright (c) Beckhoff Automation GmbH & Co. KG. All Rights Reserved.
// </copyright>
//-----------------------------------------------------------------------

namespace SE_Test
{
    using System;
    using TcHmiSrv.Core;
    using TcHmiSrv.Core.General;
    using TcHmiSrv.Core.Listeners;
    using TcHmiSrv.Core.Tools.Management;
    using ValueType = TcHmiSrv.Core.ValueType;
    using Excel = Microsoft.Office.Interop.Excel;
    using System.IO;
    using System.Runtime.InteropServices;
    using System.Data.SqlClient;
    using System.Data.SQLite;
    using System.Data;
    using System.Threading;
    using Newtonsoft.Json;


    // Represents the default type of the TwinCAT HMI server extension.
    public class SE_Test : IServerExtension
    {
        private readonly RequestListener requestListener = new RequestListener();

        private readonly Data data = new Data();

        private readonly Random rand = new Random("SE_Test".GetHashCode());

        bool tempBoolVal;
        string tempGettingDataStatus = "Server Extension (Export Excel) is ready to use";

        // Initializes the TwinCAT HMI server extension.
        public ErrorValue Init()
        {
            try
            {
                // Add event handlers
                this.requestListener.OnRequest += this.OnRequest;

                TcHmiLogger.Send(Severity.Info, "MESSAGE_INIT");
                return ErrorValue.HMI_SUCCESS;
            }
            catch (Exception ex)
            {
                TcHmiLogger.Send(Severity.Error, "ERROR_INIT", ex.ToString());
                return ErrorValue.HMI_E_EXTENSION_LOAD;
            }
        }

        // Called when a client requests a symbol from the domain of the TwinCAT HMI server extension.
        private void OnRequest(object sender, TcHmiSrv.Core.Listeners.RequestListenerEventArgs.OnRequestEventArgs e)
        {
            ErrorValue ret = ErrorValue.HMI_SUCCESS;
            Context context = e.Context;
            CommandGroup commands = e.Commands;

            try
            {
                commands.Result = ExtensionErrorValue.HmiExtSuccess;
                string mapping = string.Empty;

                foreach (Command command in commands)
                {
                    mapping = command.Mapping;

                    try
                    {
                        // Use the mapping to check which command is requested
                        switch (mapping)
                        {
                            case "RandomValue":
                                ret = this.RandomValue(command);
                                break;

                            case "MaxRandom":
                                ret = this.MaxRandom(command);
                                break;

                            case "MaxRandomFromConfig":
                                ret = this.MaxRandomFromConfig(context, command);
                                break;
                            case "BoolTestVar":
                                ret = this.BoolTestVar(command);
                                break;
                            case "StatusVar":
                                ret = this.StatusVar(command);
                                break;
                            case "GettingDataStatus":
                                ret = this.GettingDataStatus(command);
                                break;
                            case "ExportExcel":
                                ret = this.ExportExcel(command);
                                break;
                            default:
                                ret = ErrorValue.HMI_E_EXTENSION;
                                break;
                        }

                        // if (ret != ErrorValue.HMI_SUCCESS)
                        //   Do something on error
                    }
                    catch (Exception ex)
                    {
                        command.ExtensionResult = ExtensionErrorValue.HmiExtFail;
                        command.ResultString = TcHmiLogger.Localize(context, "ERROR_CALL_COMMAND", new string[] { mapping, ex.ToString() });
                    }
                }
            }
            catch (Exception ex)
            {
                commands.Result = ExtensionErrorValue.HmiExtFail;
                throw new TcHmiException(ex.ToString(), (ret == ErrorValue.HMI_SUCCESS) ? ErrorValue.HMI_E_EXTENSION : ret);
            }
        }

        // Generates a random value and writes it to the read value of the specified command.
        private ErrorValue RandomValue(Command command)
        {

            command.ReadValue = this.rand.Next(this.data.MaxRandom) + 1;

            command.ExtensionResult = ExtensionErrorValue.HmiExtSuccess;
            return ErrorValue.HMI_SUCCESS;
        }

        #region Sample code Server Extension

        // Gets or sets the maximum random value.
        private ErrorValue MaxRandom(Command command)
        {
            if ((command.WriteValue != null) && (command.WriteValue.Type == ValueType.Int32))
            {
                this.data.MaxRandom = command.WriteValue;
            }

            command.ReadValue = this.data.MaxRandom;

            command.ExtensionResult = ExtensionErrorValue.HmiExtSuccess;
            return ErrorValue.HMI_SUCCESS;
        }

        // Gets the maximum random value from the configuration of the TwinCAT HMI server extension.
        private ErrorValue MaxRandomFromConfig(Context context, Command command)
        {
            command.ReadValue = TcHmiApplication.Host.GetConfigValue(context, "MaxRandom");

            command.ExtensionResult = ExtensionErrorValue.HmiExtSuccess;
            return ErrorValue.HMI_SUCCESS;
        }

        #endregion

        #region Test function of Server Extension
        private ErrorValue BoolTestVar(Command command)
        {

            if ((command.WriteValue != null) && (command.WriteValue.Type == ValueType.Bool))
            {
                this.data.BoolTestVar = command.WriteValue;
                if (this.data.BoolTestVar == true)
                {
                    //ExportExcelFile();
                }
            }

            command.ReadValue = this.data.BoolTestVar;
            tempBoolVal = this.data.BoolTestVar;

            command.ExtensionResult = ExtensionErrorValue.HmiExtSuccess;
            return ErrorValue.HMI_SUCCESS;
        }
        private ErrorValue StatusVar (Command command)
        {
            if (tempBoolVal) command.ReadValue = "True value om BoolTestVar. OK";
            else
                command.ReadValue = "False value om BoolTestVar. OK";
            command.ExtensionResult = ExtensionErrorValue.HmiExtSuccess;
            return ErrorValue.HMI_SUCCESS;
        }
        #endregion

        //Các chương trình con đọc ghi và xuất file EXCEL

        private ErrorValue GettingDataStatus (Command command)
        {
            command.ReadValue = tempGettingDataStatus;
            return ErrorValue.HMI_SUCCESS;
        }


        private ErrorValue ExportExcel(Command command)
        {

            if ((command.WriteValue != null) && (command.WriteValue.Type == ValueType.Bool))
            {
                this.data.ExportExcel = command.WriteValue;
                if (this.data.ExportExcel == true)
                {
                    tempGettingDataStatus = "";
                    /* Phần code dành để chạy xuất file excel trên 1 thread mới*/

                    //Thread t = new Thread(() => { ExportExcelFile(); });
                    //t.IsBackground = true;
                    //t.Start();
                    // try
                    // {
                    //     ExportExcelFile();
                    // }
                    //catch (Exception e)
                    // {
                    //     //Skip error
                    // }
                    ExportExcelFile();
                }
            }
            command.ReadValue = this.data.ExportExcel;

            command.ExtensionResult = ExtensionErrorValue.HmiExtSuccess;
            return ErrorValue.HMI_SUCCESS;
        }



        //Hàm phụ trợ LẤY DỮ LIỆU TỪ DATABASE VÀ XUẤT FILE EXCEL
        private void ExportExcelFile ()
        {
            #region code excel
            //SQLite
            tempGettingDataStatus += "Creating SQlite Connection ... \n";
            //Đường link SQL theo thư mục
            //SQLiteConnection sqlite = new SQLiteConnection("Data Source=D:\\WORK SPACE\\Beckhoff Internship\\Server Extension\\Test Database\\historize.db;New=False;");
            SQLiteConnection sqlite = new SQLiteConnection("Data Source=C:\\ProgramData\\Beckhoff\\TF2000 TwinCAT 3 HMI Server\\service\\TcHmiProject\\historize.db;New=False;");
            SQLiteDataAdapter ad;
            DataTable dt = new DataTable();
            SQLiteCommand cmd;
            sqlite.Open();  //Initiate connection to the db
            cmd = sqlite.CreateCommand();

            cmd.CommandText = "SELECT t2.value, DATETIME(ROUND(t1.updated / 1000), 'unixepoch', 'localtime') AS TimeStamp FROM record AS t1 INNER JOIN value AS t2 ON t2.[id] = t1.[valueid] WHERE t1.[symbol] = 'PLC1.MAIN.Value1' AND t1.filter = 'RAW_DATA' AND t1.[updated]>=1606898939933";
            ad = new SQLiteDataAdapter(cmd);

            ad.Fill(dt);
            //Open file stream
            tempGettingDataStatus += "Creating JSON file ... \n";
            using (StreamWriter file = File.CreateText(@"C:\jsonTest\SQLiteData.json"))
            {
                JsonSerializer serializer = new JsonSerializer();
                serializer.Serialize(file, dt);
            }
            tempGettingDataStatus += "Creating JSON file is complete ... \n";

            //instantiate excel objects (application, workbook, worksheets)
            tempGettingDataStatus += "Writing Database into Excel file ... \n";
            Excel.Application XlObj = new Excel.Application();
            XlObj.Visible = false;
            XlObj.DisplayAlerts = false;
            XlObj.ScreenUpdating = false;
            if (XlObj == null)
            {
                tempGettingDataStatus += "Khong the su dung thu vien Excel \n";
            }

            //No error in this next 2 code lines.
            Excel._Workbook WbObj = (Excel.Workbook)(XlObj.Workbooks.Add(""));
            Excel._Worksheet WsObj = (Excel.Worksheet)WbObj.ActiveSheet;

            
            //run through datatable and assign cells to values of datatable
            try
            {
                int row = 1; int col = 1;
                foreach (DataColumn column in dt.Columns)
                {
                    //adding columns
                    WsObj.Cells[row, col] = column.ColumnName;
                    col++;
                }
                //reset column and row variables
                col = 1;
                row++;
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    //adding data
                    foreach (var cell in dt.Rows[i].ItemArray)
                    {
                        if (col == 2) WsObj.Cells[row, col].NumberFormat = "dd-MM-yyyy HH:mm:ss";
                        WsObj.Cells[row, col] = cell;
                        col++;
                    }
                    col = 1;
                    row++;
                }
            }
            catch (Exception e)
            {
                tempGettingDataStatus += "Error in Create data in dt WbObj: " + e.Message.ToString() + "\n";
            }

            tempGettingDataStatus += "Note: Code is running \n";

            File.Delete(@"C:\excelTest\SQLiteData.xlsx");
            bool failed = false;
            //do
            //{
            //    try
            //    {

            //        WbObj.SaveAs(@"C:\excelTest\SQLiteData.xlsx");

            //    }
            //    catch (System.Runtime.InteropServices.COMException e)
            //    {
            //        failed = true;
            //    }

            //} while (failed);        
            WbObj.SaveAs(@"C:\excelTest\SQLiteData.xlsx");
            try
            {
                WbObj.Close();
            }
            catch (Exception e)
            {
                tempGettingDataStatus += "Error in Close WbObj: " +e.Message.ToString() + "\n";
            }
           
            
            #endregion
            tempGettingDataStatus += "Opening Excel file ... \n";
            Excel.Application xl = new Excel.Application();
            xl.Visible = true;
            try
            {
                Excel.Workbook wb = xl.Workbooks.Open(@"C:\excelTest\SQLiteData.xlsx");
            }
            catch (Exception e)
            {
                tempGettingDataStatus += "Error in Open" + e.Message.ToString() + "\n";
            }
            
            xl.ActiveWindow.Activate();
            tempGettingDataStatus += "Export Excel file complete ... \n";
            sqlite.Close();
            this.data.ExportExcel = false;
        }
    }
}
