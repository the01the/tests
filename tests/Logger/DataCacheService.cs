using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using System;
using System.Collections.Generic;
using System.IO;
using System.Text.Json;
using System.Threading.Tasks;
using Microsoft.Extensions.Logging; // ILogger を使用
namespace tests.Logger
{

    public class DataCacheService
    {
        private readonly ILogger<DataCacheService> _logger;
        private readonly string _cacheDirectory;
        private const string CacheFileName = "db_cache.json";
        private const int CacheRetentionDays = 7; // キャッシュ保持期間

        public DataCacheService(ILogger<DataCacheService> logger)
        {
            _logger = logger;
            // ユーザーのローカルAppDataフォルダにアプリケーション固有のキャッシュフォルダを作成
            _cacheDirectory = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "YourAppName", "cache");

            // ディレクトリが存在しない場合は作成
            if (!Directory.Exists(_cacheDirectory))
            {
                Directory.CreateDirectory(_cacheDirectory);
                _logger.LogInformation($"キャッシュディレクトリを作成しました: {_cacheDirectory}");
            }
        }

        /// <summary>
        /// データをJSON形式でキャッシュファイルに保存します。
        /// </summary>
        /// <param name="data">キャッシュするデータリスト</param>
        public async Task SaveDataCacheAsync(List<MyDataItem> data)
        {
            string cacheFilePath = Path.Combine(_cacheDirectory, CacheFileName);

            try
            {
                var options = new JsonSerializerOptions { WriteIndented = true }; // 読みやすく整形
                string jsonString = JsonSerializer.Serialize(data, options);
                await File.WriteAllTextAsync(cacheFilePath, jsonString);
                _logger.LogInformation($"データベースキャッシュを {cacheFilePath} に保存しました。");
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "データベースキャッシュの保存中にエラーが発生しました。");
            }
        }

        /// <summary>
        /// キャッシュファイルを読み込み、有効期限内の場合はデータを返します。
        /// </summary>
        /// <returns>キャッシュデータ、または有効期限切れ/ファイルが存在しない場合はnull</returns>
        public async Task<List<MyDataItem>> LoadDataCacheAsync()
        {
            string cacheFilePath = Path.Combine(_cacheDirectory, CacheFileName);

            // 古いキャッシュファイルをクリーンアップ
            CleanOldCacheFiles(_cacheDirectory, CacheRetentionDays);

            if (File.Exists(cacheFilePath))
            {
                try
                {
                    // ファイルの最終更新日時をチェック（既にCleanOldCacheFilesで処理済みだが念のため）
                    if (DateTime.Now.Subtract(File.GetLastWriteTime(cacheFilePath)).TotalDays > CacheRetentionDays)
                    {
                        _logger.LogInformation($"キャッシュファイル {cacheFilePath} は有効期限切れです。");
                        // CleanOldCacheFiles で削除されるはずだが、もし残っていたらここで削除するロジックも追加可能
                        // File.Delete(cacheFilePath);
                        return null;
                    }

                    string jsonString = await File.ReadAllTextAsync(cacheFilePath);
                    var cachedData = JsonSerializer.Deserialize<List<MyDataItem>>(jsonString);
                    _logger.LogInformation($"データベースキャッシュを {cacheFilePath} から読み込みました。");
                    return cachedData;
                }
                catch (Exception ex)
                {
                    _logger.LogError(ex, "データベースキャッシュの読み込み中にエラーが発生しました。ファイルが破損している可能性があります。");
                    // 読み込みに失敗した場合、破損したファイルを削除することも検討
                    try
                    {
                        File.Delete(cacheFilePath);
                        _logger.LogWarning($"破損したキャッシュファイル {cacheFilePath} を削除しました。");
                    }
                    catch (Exception deleteEx)
                    {
                        _logger.LogError(deleteEx, $"破損したキャッシュファイルの削除に失敗しました: {cacheFilePath}");
                    }
                    return null;
                }
            }
            _logger.LogInformation($"キャッシュファイル {cacheFilePath} が見つかりませんでした。");
            return null; // ファイルが存在しない場合
        }

        /// <summary>
        /// 指定されたディレクトリ内の古いキャッシュファイルを削除します。
        /// </summary>
        /// <param name="directoryPath">キャッシュファイルが存在するディレクトリのパス</param>
        /// <param name="retentionDays">ファイルを保持する日数</param>
        private void CleanOldCacheFiles(string directoryPath, int retentionDays)
        {
            _logger.LogDebug($"キャッシュディレクトリ '{directoryPath}' のクリーンアップを開始します。");
            try
            {
                // ディレクトリが存在しない場合は処理をスキップ
                if (!Directory.Exists(directoryPath))
                {
                    _logger.LogDebug($"キャッシュディレクトリ '{directoryPath}' が存在しないため、クリーンアップをスキップします。");
                    return;
                }

                foreach (string filePath in Directory.EnumerateFiles(directoryPath))
                {
                    // ファイル名がCacheFileNameと一致する場合のみチェック（複数ファイルがある可能性も考慮）
                    if (Path.GetFileName(filePath).Equals(CacheFileName, StringComparison.OrdinalIgnoreCase))
                    {
                        if ((DateTime.Now - File.GetLastWriteTime(filePath)).TotalDays > retentionDays)
                        {
                            File.Delete(filePath);
                            _logger.LogInformation($"古いキャッシュファイル {filePath} を削除しました。");
                        }
                    }
                }
                _logger.LogDebug($"キャッシュディレクトリ '{directoryPath}' のクリーンアップが完了しました。");
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, $"古いキャッシュファイルのクリーンアップ中にエラーが発生しました: {directoryPath}");
            }
        }
    }
}
