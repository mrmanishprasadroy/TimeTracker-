/*
 * Author: Manish Roy
 * Purpose: Windows service to record user activity based on start, stop and power mode 
 * version : v1.01
 * Database : Microsoft Access 2010
 * 
 * 
 *                      _oo0oo_
 *                     o8888888o
 *                     88" . "88
 *                     (| -_- |)
 *                     0\  =  /0
 *                   ___/`---'\___
 *                 .' \\|     |// '.
 *                / \\|||  :  |||// \
 *               / _||||| -:- |||||- \
 *              |   | \\\  -  /// |   |
 *              | \_|  ''\---/''  |_/ |
 *              \  .-\__  '-'  ___/-. /
 *            ___'. .'  /--.--\  `. .'___
 *         ."" '<  `.___\_<|>_/___.' >' "".
 *        | | :  `- \`.;`\ _ /`;.`/ - ` : | |
 *        \  \ `_.   \_ __\ /__ _/   .-` /  /
 *    =====`-.____`.___ \_____/___.-`___.-'=====
 *                      `=---='
 * 
 * 
 *    ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
*/
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Linq;
using System.ServiceProcess;
using System.Text;
using System.Data.OleDb;
using System.Configuration;
using Microsoft.Win32;

namespace WS_SAP
{
    /// <summary>
    /// Event type defination 
    /// </summary>
    enum eventType {  Test,

                      Shutdown,

                      SleepMode,

                      Startup, 
 
                      ConsoleConnect,
    
                      ConsoleDisconnect,
    
                      RemoteConnect,
    
                      RemoteDisconnect,
    
                      SessionLock,
    
                      SessionLogoff,
  
                      SessionLogon,

                      SessionRemoteControl,
                     
                      SessionUnlock,
                      
                      PowerLogOff,
                      
                      PowerLogOn
    }
 /// <summary>
 /// services intilization 
 /// </summary>
    public partial class Service1 : ServiceBase       
    {
        private OleDbConnection bookCon;
        private OleDbCommand DbCommand = new OleDbCommand();
        private string ConnectioString;
        private  DateTime StartTime;
     
      
        /// <summary>
        /// constructor
        /// </summary>
        public Service1()
        {
            InitializeComponent();
            ConnectioString = ConfigurationManager.ConnectionStrings["SAP_Time__Tracking.Properties.Settings.testConnectionString"].ConnectionString;
            SystemEvents.PowerModeChanged += new PowerModeChangedEventHandler(SystemEvents_PowerModeChanged);
            Microsoft.Win32.SystemEvents.SessionEnded += new SessionEndedEventHandler(SystemEvents_SessionEnded);
        }
        /// <summary>
        /// session end event handler
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>

        void SystemEvents_SessionEnded(object sender, SessionEndedEventArgs e)
        {
            Sys_Action(eventType.Shutdown);
        }
        /// <summary>
        /// shutdown overloaded
        /// </summary>
        protected override void OnShutdown()
        {
            Sys_Action(eventType.Shutdown);
            base.OnShutdown();
        }
        /// <summary>
        /// session change event handler 
        /// </summary>
        /// <param name="changeDescription"></param>
        protected override void OnSessionChange(SessionChangeDescription changeDescription)
        {
           
           // System.Diagnostics.Debugger.Launch();
            switch (changeDescription.Reason)
            {
                case SessionChangeReason.ConsoleConnect:
                    Sys_Action(eventType.ConsoleConnect);
                    break;
                case SessionChangeReason.ConsoleDisconnect:
                    Sys_Action(eventType.ConsoleDisconnect);
                    break;
                case SessionChangeReason.SessionLock:
                    Sys_Action(eventType.SessionLock);
                    break;
                case SessionChangeReason.SessionLogoff:
                    Sys_Action(eventType.SessionLogoff);
                    break;
                case SessionChangeReason.SessionLogon:
                    Sys_Action(eventType.SessionLogon);
                    break;
                case SessionChangeReason.SessionUnlock:
                    Sys_Action(eventType.SessionUnlock);
                    break;
            }
            
           // base.OnSessionChange(changeDescription);
        }
       /// <summary>
       /// data base transaction implementation 
       /// </summary>
       /// <param name="a"></param>
        void Sys_Action(eventType a)
        {
            try
            {
                //select
               // System.Diagnostics.Debugger.Launch();
                string T_dur = string.Empty;
                string Time = DateTime.Now.ToString();
                int result = -1;
                int SAPID = 0;
                DataTable dataTable = new DataTable();
                DataSet ds = new DataSet();
               
              
                OleDbDataAdapter dAdapter = new OleDbDataAdapter("select * from SAP_TIME_RECORD where ID = (select max(ID) from SAP_TIME_RECORD) ", ConnectioString);
                OleDbCommandBuilder cBuilder = new OleDbCommandBuilder(dAdapter);
                dAdapter.Fill(dataTable);
             
                ds.Tables.Add(dataTable);
                if (ds.Tables[0].Rows.Count > 0 && string.IsNullOrEmpty(Convert.ToString(ds.Tables[0].Rows[0]["Duration"])))
                {
                    StartTime = Convert.ToDateTime(ds.Tables[0].Rows[0]["StartTime"]);
                    SAPID = Convert.ToInt32(ds.Tables[0].Rows[0]["ID"]); 
                    TimeSpan Duartion = DateTime.Parse(DateTime.Now.ToString()).Subtract(DateTime.Parse(StartTime.ToString()));

                    bookCon = new OleDbConnection(ConnectioString);
                    bookCon.Open();
                    DbCommand.Connection = bookCon;
                    DbCommand.CommandText = "update  SAP_TIME_RECORD set EndTime = '" + Time + "' , Duration = '" + Duartion.ToString() + "'  where ID =" + SAPID;
                    result = DbCommand.ExecuteNonQuery();
                    bookCon.Dispose();
                }

                if (a != eventType.Shutdown)
                {
                    bookCon = new OleDbConnection(ConnectioString);
                    bookCon.Open();
                    DbCommand.Connection = bookCon;
                    DbCommand.CommandText = "insert into SAP_TIME_RECORD (StartTime,To_Date,Event) values ('" + Time + "','" + DateTime.Now.Date.ToString() + "','" + a.ToString() + "')";

                    result = DbCommand.ExecuteNonQuery();
                    bookCon.Dispose();
                }
                
            }
            catch (Exception e)
            {

                bookCon.Dispose();
            }
        }
        /// <summary>
        /// power mode chage handler
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>

       void SystemEvents_PowerModeChanged(object sender, Microsoft.Win32.PowerModeChangedEventArgs e)
        {
            //System.Diagnostics.Debugger.Launch();
            switch (e.Mode)
            {
                case PowerModes.Resume:
                     Sys_Action(eventType.PowerLogOn);
                    break;
                case PowerModes.StatusChange:
                    break;
                case PowerModes.Suspend:
                    Sys_Action(eventType.PowerLogOff);
                    break;
            }

            
        }
        /// <summary>
        /// on start 
        /// </summary>
        /// <param name="args"></param>
        protected override void OnStart(string[] args)
        {
            
            Sys_Action(eventType.Startup);
           
        }
        /// <summary>
        /// on service stop 
        /// </summary>
        protected override void OnStop()
        {
            Sys_Action(eventType.Shutdown);
        }
    }
}
