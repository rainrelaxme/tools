using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Threading.Tasks;
using InfluxData.Net.Common.Enums;
using InfluxData.Net.InfluxDb;
using InfluxData.Net.InfluxDb.Models;
using InfluxData.Net.InfluxDb.Models.Responses;

namespace weldData_collection
{
    public class InfluxDataNetService
    {
        private readonly IInfluxDbClient _client;

        public InfluxDataNetService(string url, string username, string password, string database)
        {
            _client = new InfluxDbClient(
                url,
                username,
                password,
                InfluxDbVersion.Latest,
                queryLocation: QueryLocation.FormData);
        }

        public async Task TestConnectionAsync()
        {
            try
            {
                var ping = await _client.Diagnostics.PingAsync();
                Console.WriteLine($"Ping 响应时间: {ping.ResponseTime}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"连接测试失败: {ex.Message}");
                throw;
            }
        }

        public async Task EnsureDatabaseAsync(string database)
        {
            var dbs = await _client.Database.GetDatabasesAsync();
            if (dbs.All(d => d.Name != database))
            {
                await _client.Database.CreateDatabaseAsync(database);
                Console.WriteLine($"已创建数据库: {database}");
            }
        }

        public async Task<int> WriteRowsAsync(IEnumerable<WeldCsvRecord> rows, string database, string machineNumber, string measurement)
        {
            var latestTime = await GetLatestTimestampAsync(database, measurement);

            var parsedRows = rows
                .Select(r => new { Record = r, ParsedTime = TryParseTime(r.Endtime) })
                .Where(x => x.ParsedTime.HasValue)
                .Where(x => x.ParsedTime.Value > latestTime)
                .OrderBy(x => x.ParsedTime.Value) // 按时间顺序写入
                .ToList();

            if (!parsedRows.Any())
            {
                Console.WriteLine("没有新增数据需要写入（时间未超过已存在的最新 Endtime）。");
                return 0;
            }

            var points = parsedRows.Select(x => BuildPoint(x.Record, measurement, x.ParsedTime!.Value, machineNumber)).ToList();
            
            if (points.Count > 0)
            {
                try
                {
                    await _client.Client.WriteAsync(points, database);
                    Console.WriteLine($"已按 Endtime 过滤，写入 {points.Count} 行。");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"写入 InfluxDB 时发生错误: {ex.Message}");
                    throw;
                }
            }
            
            return points.Count;
        }

        private async Task<DateTime> GetLatestTimestampAsync(string database, string measurement)
        {
            var query = $"select * from {measurement} order by time desc limit 1";
            var series = await _client.Client.QueryAsync(query, database);
            var latest = DateTime.MinValue;

            if (series == null)
            {
                return latest;
            }

            foreach (var serie in series)
            {
                if (serie?.Values == null)
                {
                    continue;
                }

                foreach (var row in serie.Values)
                {
                    // Influx 返回的第一列为时间字符串
                    if (row.Count > 0 && row[0] != null && DateTime.TryParse(row[0].ToString(), out var t))
                    {
                        // 统一为 UTC
                        latest = t.ToUniversalTime();
                        return latest;
                    }
                }
            }

            return latest;
        }

        private static Point BuildPoint(WeldCsvRecord record, string measurement, DateTime timestampUtc, string machineNumber)
        {
            var tags = new Dictionary<string, object>
            {
                { "sn", record.SN },
                { "machineNo", machineNumber }
            };

            return new Point
            {
                Name = measurement,
                Tags = tags,
                Fields = new Dictionary<string, object>
                {
                    { "pointCount", ParseInt(record.PointCount) },
                    { "temp1", ParseDouble(record.Temp1) },
                    { "dura1", ParseDouble(record.Dura1) },
                    { "len1", ParseDouble(record.Len1) },
                    { "speed1", ParseDouble(record.Speed1) },
                    { "temp2", ParseDouble(record.Temp2) },
                    { "dura2", ParseDouble(record.Dura2) },
                    { "len2", ParseDouble(record.Len2) },
                    { "speed2", ParseDouble(record.Speed2) },
                    { "temp3", ParseDouble(record.Temp3) },
                    { "dura3", ParseDouble(record.Dura3) },
                    { "len3", ParseDouble(record.Len3) },
                    { "speed3", ParseDouble(record.Speed3) },
                    { "temp4", ParseDouble(record.Temp4) },
                    { "dura4", ParseDouble(record.Dura4) },
                    { "len4", ParseDouble(record.Len4) },
                    { "speed4", ParseDouble(record.Speed4) }
                },
                Timestamp = timestampUtc
            };
        }

        private static int ParseInt(string value)
        {
            return int.TryParse(value, NumberStyles.Any, CultureInfo.InvariantCulture, out var result)
                ? result
                : 0;
        }

        private static double ParseDouble(string value)
        {
            if (double.TryParse(value, NumberStyles.Any, CultureInfo.InvariantCulture, out var result))
            {
                return result;
            }

            if (double.TryParse(value, NumberStyles.Any, CultureInfo.GetCultureInfo("zh-CN"), out result))
            {
                return result;
            }

            return 0d;
        }

        private static DateTime? TryParseTime(string value)
        {
            if (string.IsNullOrWhiteSpace(value))
            {
                return null;
            }

            if (DateTime.TryParse(value, CultureInfo.InvariantCulture, DateTimeStyles.AssumeUniversal | DateTimeStyles.AdjustToUniversal, out var time))
            {
                return time.ToUniversalTime();
            }

            if (DateTime.TryParse(value, CultureInfo.GetCultureInfo("zh-CN"), DateTimeStyles.AssumeLocal, out time))
            {
                return time.ToUniversalTime();
            }

            return null;
        }
    }
}
