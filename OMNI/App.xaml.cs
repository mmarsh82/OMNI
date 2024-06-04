using Microsoft.Win32;
using System;
using System.Data;
using System.Data.SqlClient;
using System.Deployment.Application;
using System.Diagnostics;
using System.Reflection;
using System.Windows;
using System.Windows.Threading;

namespace OMNI
{
    /// <summary>
    /// Interaction logic for App.xaml
    /// </summary>
    public partial class App : Application
    {
        #region Properties

        /// <summary>
        /// Checks to see if there are any available updates for OMNI
        /// </summary>
        public static bool IsUpdateAvailable
        {
            get
            {
                try
                {
                    if (SqlConAsync.State == ConnectionState.Open)
                    {
                        using (var cmd = new SqlCommand(@"USE [OMNI];
                                                          SELECT * FROM [version];", SqlConAsync))
                        {
                            if (cmd.ExecuteScalar().ToString().Equals(Assembly.GetExecutingAssembly().GetName().Version.ToString()))
                            {
                                return false;
                            }
                            return true;
                        }
                    }
                    return false;
                }
                catch (Exception)
                {
                    return false;
                }
            }
        }

        /// <summary>
        /// Is Microsoft Word installed on the host machine
        /// </summary>
        public static bool IsWordInstalled { get; private set; }

        /// <summary>
        /// OMNI SQL Async Connection
        /// </summary>
        public static SqlConnection SqlConAsync { get; private set; }

        /// <summary>
        /// OMNI Sql DataBase name to read and write to
        /// </summary>
        public static string DataBase { get; set; }
        public static bool IsExiting { get; set; }

        #endregion

        /// <summary>
        /// Update OMNI to the latest version
        /// </summary>
        public static void Update()
        {
            {
                ApplicationDeployment.CurrentDeployment.Update();
                Current.Shutdown();
                Process.Start(OMNI.Properties.Settings.Default.omniAppLocation);
            }
        }

        /// <summary>
        /// Manual Update for OMNI developers
        /// </summary>
        public static void ManualUpdate()
        {
            try
            {
                if (ApplicationDeployment.IsNetworkDeployed)
                {
                    var ad = ApplicationDeployment.CurrentDeployment;
                    if (ad.CheckForUpdate())
                    {
                        var uc = ad.CheckForDetailedUpdate();
                        var s = uc.IsUpdateRequired ? "required. Update OMNI now for the best performance" : "not required. No further action is required.";
                        MessageBox.Show($@"OMNI last checked for updates on {ad.TimeOfLastUpdateCheck.ToLongDateString()}.
                                           The current minimum required version number is {uc.MinimumRequiredVersion}.
                                           There is an available update version number {uc.AvailableVersion} that has not been applied.
                                           The available update is currently marked as {s}", "Manual update results", MessageBoxButton.OK, MessageBoxImage.Information, MessageBoxResult.OK);
                    }
                    else
                    {
                        MessageBox.Show("The check has completed successfully.\nThere is no pending updates for OMNI.", "Manual Update Check", MessageBoxButton.OK, MessageBoxImage.Information, MessageBoxResult.OK);
                    }
                }
                else
                {
                    MessageBox.Show("OMNI is currently not networked deployed.", "Incorrect Deployment", MessageBoxButton.OK, MessageBoxImage.Exclamation, MessageBoxResult.OK);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($@"An exception was thrown during the manual update check.
                                   Source: {ex.Source}
                                   Default Message: {ex.Message}
                                   Please review the below stack trace;
                                   [{ex.StackTrace}", "Update exception", MessageBoxButton.OK, MessageBoxImage.Error, MessageBoxResult.OK);
            }
        }

        /// <summary>
        /// Current Application Constructor
        /// </summary>
        public App()
        {
            IsExiting = false;
            try
            {
                using (var regWord = Registry.ClassesRoot.OpenSubKey("Word.Application"))
                {
                    IsWordInstalled = regWord != null;
                }
            }
            catch (Exception)
            {
                IsWordInstalled = false;
            }
            try
            {
                if (SqlConAsync == null)
                {
                    SqlConAsync = new SqlConnection(OMNI.Properties.Settings.Default.SqlConnectionString);
                    SqlConAsync.OpenAsync();
                }

            }
            catch (Exception)
            {
                MessageBox.Show("OMNI heavily relies on the ability to connect with it's database.\nThere is currently no available connection.\nMany of OMNI's functions will be disabled.", "Connection Failed", MessageBoxButton.OK, MessageBoxImage.Warning, MessageBoxResult.OK);
            }
            Current.Exit += App_Exit;
            AppDomain.CurrentDomain.UnhandledException += App_ExceptionCrash;
            Current.DispatcherUnhandledException += App_DispatherCrash;
            SystemEvents.PowerModeChanged += OnPowerChange;
            Helpers.UpdateTimer.IntializeUpdateTimer(new TimeSpan(0,0,10));
            SqlConAsync.StateChange += AppSqlCon_StateChangeAsync;
            DataBase = "OMNI";
        }

        /// <summary>
        /// Application On Startup method for running cmd input overrides
        /// </summary>
        /// <param name="e">start up events sent from the application.exe</param>
        protected override void OnStartup(StartupEventArgs e)
        {
            string[] startUpArgs = null;
            try
            {
                if (AppDomain.CurrentDomain.SetupInformation.ActivationArguments != null)
                {
                    startUpArgs = AppDomain.CurrentDomain.SetupInformation.ActivationArguments.ActivationData;
                }
            }
            catch (NullReferenceException)
            {
                startUpArgs = e.Args;
            }
            if (startUpArgs != null)
            {
                foreach (string s in startUpArgs)
                {
                    switch (s.Remove(0,1))
                    {
                        case "t":
                            //TrainingStatus = true;
                            break;
                        case "m":
                            Helpers.MapForm.PdfForFields();
                            Current.Shutdown();
                            break;
                    }
                }
            }
        }

        /// <summary>
        /// SqlConnection state change watch
        /// </summary>
        /// <param name="sender">empty object</param>
        /// <param name="e">Connection State Change Events</param>
        private async static void AppSqlCon_StateChangeAsync(object sender, StateChangeEventArgs e)
        {
            var count = 0;
            while (!IsExiting && (SqlConAsync.State == ConnectionState.Broken || SqlConAsync.State == ConnectionState.Closed) && count <= 8)
            {
                await SqlConAsync.OpenAsync();
                count++;
            }
        }

        /// <summary>
        /// Application Exit Event
        /// </summary>
        /// <param name="sender">Current Application</param>
        /// <param name="e">Exit Event Arguments</param>
        private void App_Exit(object sender, ExitEventArgs e)
        {
            IsExiting = true;
            if (SqlConAsync != null)
            {
                SqlConAsync.Close();
                SqlConAsync.Dispose();
                SqlConAsync = null;
            }
        }

        /// <summary>
        /// Application UI Thread Exception Handler
        /// </summary>
        /// <param name="sender">Current Application</param>
        /// <param name="e">Dispatcher Thread Unhandled Exception</param>
        private static void App_DispatherCrash(object sender, DispatcherUnhandledExceptionEventArgs e)
        {
            try
            {
                e.Handled = true;
                if (SqlConAsync != null && SqlConAsync.State == ConnectionState.Open)
                {
                    Models.OMNIException.SendtoLogAsync(e.Exception.Source, e.Exception.StackTrace, e.Exception.Message, nameof(App_DispatherCrash));
                }
            }
            finally
            {
                IsExiting = true;
                if (SqlConAsync != null)
                {
                    SqlConAsync.Close();
                    SqlConAsync.Dispose();
                    SqlConAsync = null;
                }
                Current.Shutdown();
            }
        }

        /// <summary>
        /// Application Thread Domain Unhandled Exception Handler
        /// </summary>
        /// <param name="sender">Current Application Domain</param>
        /// <param name="e">Unhanded Exception</param>
        private static void App_ExceptionCrash(object sender, UnhandledExceptionEventArgs e)
        {
            if (e.IsTerminating)
            {
                try
                {
                    if (SqlConAsync != null && SqlConAsync.State == ConnectionState.Open)
                    {
                        var ex = (Exception)e.ExceptionObject;
                        Models.OMNIException.SendtoLogAsync(ex.Source, ex.StackTrace, ex.Message, nameof(App_ExceptionCrash));
                    }
                }
                finally
                {
                    IsExiting = true;
                    if (SqlConAsync != null)
                    {
                        SqlConAsync.Close();
                        SqlConAsync.Dispose();
                        SqlConAsync = null;
                    }
                    Current.Shutdown();
                }
            }
        }

        /// <summary>
        /// Called when the application PC switches power states
        /// </summary>
        /// <param name="sender">Current Application</param>
        /// <param name="e">Power Change</param>
        private void OnPowerChange(object sender, PowerModeChangedEventArgs e)
        {
            switch(e.Mode)
            {
                case PowerModes.Suspend:
                    Helpers.UpdateTimer.UpdateDispatchTimer.Stop();
                    break;
                case PowerModes.Resume:
                    Helpers.UpdateTimer.UpdateDispatchTimer.Start();
                    break;
            }
        }
    }
}
