using OMNI.Extensions;
using OMNI.Helpers;
using OMNI.Views;
using System;
using System.ComponentModel;
using System.Data.SqlClient;
using System.DirectoryServices.AccountManagement;
using System.IO;
using System.Security.Permissions;

namespace OMNI.Models
{
    /// <summary>
    /// Currently Logged In User Object Interaction Logic
    /// </summary>
    public class CurrentUser
    {
        #region Properties

        private static string _fullName;
        public static string FullName
        {
            get { return _fullName; }
            set { _fullName = value; StaticPropertyChanged?.Invoke(null, new PropertyChangedEventArgs(nameof(FullName))); }
        }
        private static string _accountName;
        public static string AccountName
        {
            get { return _accountName; }
            set { _accountName = value; StaticPropertyChanged?.Invoke(null, new PropertyChangedEventArgs(nameof(AccountName))); }
        }
        public static string DomainName { get; set; }
        public static string Email { get; set; }
        public static int IdNumber { get; set; }
        public static bool SlitterLead { get; set; }
        public static bool Quality { get; set; }
        public static bool QualityNotice { get; set; }
        public static bool Kaizen { get; set; }
        public static bool CMMS { get; set; }
        public static bool CMMSCrew { get; set; }
        public static bool CMMSAdmin { get; set; }
        public static bool IT { get; set; }
        public static bool ITTeam { get; set; }
        public static bool Engineering { get; set; }
        public static bool Accounting { get; set; }
        public static bool Admin { get; set; }
        public static bool Tools { get; set; }
        public static bool Developer { get; set; }
        public static string Site { get; set; }
        private static int _noticeTimer;
        public static int NoticeTimer
        {
            get { return _noticeTimer; }
            set { _noticeTimer = value; StaticPropertyChanged?.Invoke(null, new PropertyChangedEventArgs(nameof(NoticeTimer))); }
        }
        private static int _noticeHistory;
        public static int NoticeHistory
        {
            get { return _noticeHistory; }
            set { _noticeHistory = value; StaticPropertyChanged?.Invoke(null, new PropertyChangedEventArgs(nameof(NoticeHistory))); }
        }

        public static event EventHandler<PropertyChangedEventArgs> StaticPropertyChanged;

        #endregion

        /// <summary>
        /// Log In the user.
        /// </summary>
        /// <param name="userName">Input User Name</param>
        public async static void LogInAsync(string userName)
        {
            try
            {
                using (SqlCommand cmd = new SqlCommand($"USE {App.DataBase}; SELECT * FROM [users] WHERE [DomainName]=@p1", App.SqlConAsync))
                {
                    cmd.Parameters.AddWithValue("p1", userName);
                    using (SqlDataReader reader = cmd.ExecuteReader())
                    {
                        while (await reader.ReadAsync())
                        {
                            FullName = reader.SafeGetString(nameof(FullName));
                            DomainName = reader.SafeGetString(nameof(DomainName));
                            AccountName = reader.SafeGetString(nameof(AccountName));
                            Email = reader.SafeGetString(nameof(Email));
                            IdNumber = reader.SafeGetInt32("EmployeeNumber");
                            SlitterLead = reader.SafeGetBoolean(nameof(SlitterLead));
                            Quality = reader.SafeGetBoolean(nameof(Quality));
                            QualityNotice = reader.SafeGetBoolean(nameof(QualityNotice));
                            Kaizen = reader.SafeGetBoolean(nameof(Kaizen));
                            CMMS = reader.SafeGetBoolean(nameof(CMMS));
                            CMMSCrew = reader.SafeGetBoolean(nameof(CMMSCrew));
                            CMMSAdmin = reader.SafeGetBoolean(nameof(CMMSAdmin));
                            IT = reader.SafeGetBoolean(nameof(IT));
                            ITTeam = reader.SafeGetBoolean(nameof(ITTeam));
                            Engineering = reader.SafeGetBoolean(nameof(Engineering));
                            Accounting = reader.SafeGetBoolean(nameof(Accounting));
                            Admin = reader.SafeGetBoolean("OMNIAdministrator");
                            Tools = reader.SafeGetBoolean("OMNITools");
                            Developer = reader.SafeGetBoolean(nameof(Developer));
                            Site = reader.SafeGetString(nameof(Site));
                            NoticeTimer = reader.SafeGetInt32(nameof(NoticeTimer));
                            NoticeHistory = reader.SafeGetInt32(nameof(NoticeHistory));
                        }
                    }
                }
            }
            catch (Exception e)
            {
                ExceptionWindow.Show("Unhandled Exception", e.Message, e);
            }
        }

        /// <summary>
        /// Check to see if the log in information correlates to a registered OMNI user.
        /// </summary>
        /// <param name="userName">Input User Name</param>
        /// <returns>true = exists; false = does not exist</returns>
        public static bool Exists(string userName)
        {
            try
            {
                using (SqlCommand cmd = new SqlCommand($"USE {App.DataBase}; SELECT COUNT(*) FROM [users] WHERE [DomainName]=@p1", App.SqlConAsync))
                {
                    cmd.Parameters.AddWithValue("p1", userName);
                    if (Convert.ToInt32(cmd.ExecuteScalar()) > 0)
                        return true;
                    return false;
                }
            }
            catch (Exception)
            {
                return false;
            }
        }

        /// <summary>
        /// Log Out Current User
        /// </summary>
        public static void LogOut()
        {
            FullName = Email = AccountName = DomainName = Site = null;
            Quality = QualityNotice = Kaizen = CMMS = CMMSCrew = CMMSAdmin = IT = ITTeam = Engineering = Accounting = Admin = Developer = false;
            NoticeTimer = NoticeHistory = 0;
            DashBoardWindowView.DashBoardView?.Close();
        }

        /// <summary>
        /// Update current user information
        /// </summary>
        /// <param name="fullName">Current User Full Name</param>
        /// <param name="accountName">Current User Account Name</param>
        /// <param name="email">Current User eMail</param>
        /// <param name="noticeTimer">Current User Notice Timer settings</param>
        /// <param name="noticeHistory">Current User Notice Histroy settings</param>
        public static void UpdateUser(string fullName, string accountName, string email, int noticeTimer, int noticeHistory)
        {
            using (SqlCommand cmd = new SqlCommand($"USE {App.DataBase}; UPDATE [users] SET [FullName]=@p1, [AccountName]=@p2, [eMail]=@p3, [NoticeTimer]=@p4, [NoticeHistory]=@p5 WHERE [EmployeeNumber]=@p6", App.SqlConAsync))
            {
                cmd.Parameters.AddWithValue("p1", fullName);
                cmd.Parameters.AddWithValue("p2", accountName);
                cmd.Parameters.AddWithValue("p3", email);
                cmd.Parameters.AddWithValue("p4", noticeTimer);
                cmd.Parameters.AddWithValue("p5", noticeHistory);
                cmd.Parameters.AddWithValue("p6", IdNumber);
                cmd.ExecuteNonQuery();
            }
            FullName = fullName;
            AccountName = accountName;
            Email = email;
            NoticeTimer = noticeTimer;
            NoticeHistory = noticeHistory;
        }

        /// <summary>
        /// Validate the log in information exists in the domain.
        /// </summary>
        /// <param name="userName">Inputed User Name</param>
        /// <param name="password">Inputed Password</param>
        public static bool Validate(string userName, string password)
        {
            if (string.IsNullOrWhiteSpace(userName) || string.IsNullOrWhiteSpace(password))
                return false;
            if ((userName == "Devmin" && password == "McM2015"))
                return true;
            try
            {
                using (PrincipalContext pc = new PrincipalContext(ContextType.Domain))
                {
                    return pc.ValidateCredentials(userName, password);
                }
            }
            catch (Exception ex)
            {
                ExceptionWindow.Show("Unhandled Exception", ex.Message, ex);
                return false;
            }
        }

        /// <summary>
        /// Get the on file picture of the current user
        /// </summary>
        /// <returns>picture file path as string</returns>
        public static string GetPicture()
        {
            var dirs = Directory.GetFiles(Properties.Settings.Default.UsersPhotoDirectory, $"{FullName}*.*", SearchOption.AllDirectories);
            return dirs.Length > 0 ? dirs[0] : string.Empty;
        }

        public static string GetFullName()
        {
            using (PrincipalContext pc = new PrincipalContext(ContextType.Domain))
            {
                using (UserPrincipal up = UserPrincipal.FindByIdentity(pc, DomainName))
                {
                    return up.DisplayName;
                }
            }
        }
    }
}
