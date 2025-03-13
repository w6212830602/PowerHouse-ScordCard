using ScoreCard.Models;
using System.Collections.ObjectModel;

namespace ScoreCard.Services
{
    public interface IExportService
    {
        /// <summary>
        /// 將數據匯出為Excel檔案
        /// </summary>
        /// <param name="data">要匯出的數據</param>
        /// <param name="fileName">檔案名稱</param>
        /// <param name="title">報表標題</param>
        /// <returns>是否成功匯出</returns>
        Task<bool> ExportToExcelAsync<T>(IEnumerable<T> data, string fileName, string title);

        /// <summary>
        /// 將數據匯出為PDF檔案
        /// </summary>
        /// <param name="data">要匯出的數據</param>
        /// <param name="fileName">檔案名稱</param>
        /// <param name="title">報表標題</param>
        /// <returns>是否成功匯出</returns>
        Task<bool> ExportToPdfAsync<T>(IEnumerable<T> data, string fileName, string title);

        /// <summary>
        /// 將數據匯出為CSV檔案
        /// </summary>
        /// <param name="data">要匯出的數據</param>
        /// <param name="fileName">檔案名稱</param>
        /// <returns>是否成功匯出</returns>
        Task<bool> ExportToCsvAsync<T>(IEnumerable<T> data, string fileName);

        /// <summary>
        /// 建立並顯示銷售分析報表的打印預覽
        /// </summary>
        /// <param name="data">要打印的數據</param>
        /// <param name="title">報表標題</param>
        /// <returns>是否成功建立打印預覽</returns>
        Task<bool> PrintReportAsync<T>(IEnumerable<T> data, string title);

        /// <summary>
        /// 檢查並創建導出目錄
        /// </summary>
        /// <returns>導出目錄的路徑</returns>
        string GetExportDirectory();
    }
}