using System;
using System.Threading.Tasks;
using StackExchange.Redis;

namespace weldData_collection
{
    /// <summary>
    /// Redis 服务类，用于管理 CSV 文件处理进度
    /// 使用两个 key：latest_date（存储“最新处理的 CSV 文件名”，例如 25-12-18.csv）和 latest_row（存储行号）
    /// </summary>
    public class RedisService : IDisposable
    {
        private readonly IDatabase _db;
        private readonly ConnectionMultiplexer _redis;
        private const string LatestDateKey = "newwave_weld_latest_date_station01"; // 存储最新处理的 CSV 文件名
        private const string LatestRowKey = "newwave_weld_latest_row_station01";

        public RedisService(string connectionString, int database = 0)
        {
            _redis = ConnectionMultiplexer.Connect(connectionString);
            _db = _redis.GetDatabase(database);
        }

        /// <summary>
        /// 获取上次处理的 CSV 文件名和行号
        /// </summary>
        public async Task<(string? lastFileName, int lastRow)> GetLastProcessedInfoAsync()
        {
            var fileNameStr = await _db.StringGetAsync(LatestDateKey);
            var rowStr = await _db.StringGetAsync(LatestRowKey);

            var lastRow = 0;
            if (rowStr.HasValue && int.TryParse(rowStr, out var parsedRow))
            {
                lastRow = parsedRow;
            }

            return (fileNameStr.HasValue ? fileNameStr.ToString() : null, lastRow);
        }

        /// <summary>
        /// 更新最后处理的 CSV 文件名和行号
        /// </summary>
        public async Task UpdateLastProcessedInfoAsync(string fileName, int lastRow)
        {
            await _db.StringSetAsync(LatestDateKey, fileName);
            await _db.StringSetAsync(LatestRowKey, lastRow.ToString());
        }

        public void Dispose()
        {
            _redis?.Dispose();
        }
    }
}

