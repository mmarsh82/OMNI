using Microsoft.Office.Interop.Word;
using OMNI.Extensions;
using OMNI.Helpers;
using OMNI.Models;
using OMNI.QMS.Enumeration;
using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.DirectoryServices.ActiveDirectory;
using System.Threading.Tasks;

namespace OMNI.QMS.Model
{
    public class NCM
    {
        #region Properties

        public int Code { get; set; }
        public string Summary { get; set; }
        public string ChartCode { get; set; }
        public int Data { get; set; }

        #endregion

        /// <summary>
        /// List of NCM Codes and summaries
        /// </summary>
        /// <returns>Generated List of NCM Codes and Summaries</returns>
        public async static Task<List<NCM>> GetNCMListAsync()
        {
            try
            {
                var _ncmList = new List<NCM>();
                using (SqlCommand cmd = new SqlCommand($@"USE {App.DataBase};
                                                        SELECT * FROM [Qir_RootCategory]", App.SqlConAsync))
                {
                    using (SqlDataReader reader = cmd.ExecuteReader())
                    {
                        while (await reader.ReadAsync())
                        {
                            _ncmList.Add(new NCM { Code = reader.SafeGetInt32("CategoryID"), Summary = reader.SafeGetString("CategoryDescription") });
                        }
                    }
                }
                return _ncmList;
            }
            catch (Exception ex)
            {
                ExceptionWindow.Show("Unhandled Exception", ex.Message, ex);
                return null;
            }
        }

        public static bool IsLegacyCode(int code)
        {
            try
            {
                using (SqlCommand cmd = new SqlCommand($@"USE {App.DataBase};
                                                        SELECT COUNT([NCMCode]) FROM [ncm] WHERE [NCMCode] = @p1", App.SqlConAsync))
                {
                    cmd.Parameters.AddWithValue("p1", code);
                    if (int.TryParse(cmd.ExecuteScalar().ToString(), out int i))
                    {
                        return i > 0;
                    }
                    return false;
                }
            }
            catch (Exception ex)
            {
                ExceptionWindow.Show("Unhandled Exception", ex.Message, ex);
                return false;
            }
        }

        public static string GetLegacySummary(int code)
        {
            try
            {
                using (SqlCommand cmd = new SqlCommand($@"USE {App.DataBase};
                                                        SELECT [Summary] FROM [ncm] WHERE [NCMCode] = @p1", App.SqlConAsync))
                {
                    cmd.Parameters.AddWithValue("p1", code);
                    return cmd.ExecuteScalar().ToString();
                }
            }
            catch (Exception ex)
            {
                ExceptionWindow.Show("Unhandled Exception", ex.Message, ex);
                return null;
            }
        }

        /// <summary>
        /// List of NCM Codes and summaries
        /// </summary>
        /// <param name="workCenterNumber">Work Center Number to use for a specific list</param>
        /// <returns>Generated List of NCM Codes and Summaries</returns>
        public async static Task<List<NCM>> GetNCMListAsync(int? workCenterNumber)
        {
            var _listType = WorkCenter.GetEZType(workCenterNumber);
            if (string.IsNullOrEmpty(_listType) || _listType == "None")
            {
                return null;
            }
            var _ncmList = new List<NCM>();
            try
            {
                using (SqlCommand cmd = new SqlCommand($@"USE {App.DataBase};
                                                        SELECT * FROM [ncm] WHERE [{_listType}]=@p1", App.SqlConAsync))
                {
                    cmd.Parameters.AddWithValue("p1", 1);
                    using (SqlDataReader reader = cmd.ExecuteReader())
                    {
                        while (await reader.ReadAsync())
                        {
                            _ncmList.Add(new NCM { Code = reader.SafeGetInt32("NCMCode"), Summary = reader.SafeGetString(nameof(Summary)) });
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                ExceptionWindow.Show("Unhandled Exception", ex.Message, ex);
            }
            return _ncmList;
        }
    }
}
