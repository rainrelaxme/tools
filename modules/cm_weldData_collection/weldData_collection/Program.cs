using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using CsvHelper;
using CsvHelper.Configuration;
using CsvHelper.Configuration.Attributes;
using weldData_collection;

namespace weldData_collection
{
    internal class Program
    {
        static async Task Main()
        {
            // 注册编码提供程序（必须放在程序开始处）
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            Console.WriteLine("****************Start to scan MESDATA folder****************");

            // 根目录：MESDATA（按需修改实际路径）
            var mesRootDir = @"D:\Code\data\MESDATA";   // test地址
            //var mesRootDir = @"\\10.136.22.38\MESData";

            // 尝试连接网络共享，如果失败则使用凭据
            if (!NetworkShareService.EnsureNetworkShareConnected(mesRootDir, "mes", "cmmes"))
            {
                Console.WriteLine($"无法连接到网络共享: {mesRootDir}");
                return;
            }

            // InfluxDB 配置（根据实际环境修改）
            var influxUrl = "http://localhost:8086";    //test地址
            //var influxUrl = "http://47.102.85.235:8086";
            var influxUser = "root";
            var influxPassword = "root";
            var database = "all_collect_v_3";
            var table = "NewWave_Weld_Station01";

            // Redis 配置（根据实际环境修改）
            var redisHost = "127.0.0.1:6379";   // test地址
            //var redisConnectionString = "10.136.11.82:6379";   // test地址
            //var redisHost = "47.102.85.235:6379";
            //var redisPassword = "cmmes1234!@#$";
            var redisPassword = "123456";
            var redisConnectionString = $"{redisHost},password={redisPassword}";
            var redisDatabase = 0; // Redis 数据库编号

            // 机器编号配置（固定参数值）
            var machineNumber = "M001"; // 根据实际机器编号修改

            try
            {
                var influxService = new InfluxDataNetService(influxUrl, influxUser, influxPassword, database);
                await influxService.EnsureDatabaseAsync(database);

                using var redisService = new RedisService(redisConnectionString, redisDatabase);
                var totalInserted = await ProcessMesDirectoryAsync(mesRootDir, influxService, redisService, database, machineNumber, table);

                Console.WriteLine($"所有文件处理完成，本次共新增写入 {totalInserted} 行数据到 InfluxDB.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"程序执行出错: {ex.Message}");
            }
        }

        /// <summary>
        /// 扫描 MESDATA 目录：
        /// - 第一层目录为 MESDATA
        /// - 第二层子文件夹为月份（如 2025-12）
        /// - 第三层为具体 CSV 文件（如 25-12-18.csv）
        /// - 逐个文件读取并写入新增数据
        /// - 使用 Redis 记录最后处理的 CSV 文件名和行号，下次只处理该文件之后及该文件内新增的行
        /// </summary>
        private static async Task<int> ProcessMesDirectoryAsync(string mesRootDir, InfluxDataNetService influxService, RedisService redisService, string database, string machineNumber, string measurement)
        {
            if (!Directory.Exists(mesRootDir))
            {
                Console.WriteLine($"MES 目录不存在: {mesRootDir}");
                return 0;
            }

            var totalInserted = 0;

            // 从 Redis 获取上次处理的 CSV 文件名和行号
            var (lastFileName, lastRow) = await redisService.GetLastProcessedInfoAsync();
            Console.WriteLine($"Redis 记录：latest_date(文件名)={lastFileName ?? "无"}, latest_row={lastRow}");

            // 第二层：月份目录（如 2025-12），按名称排序
            var monthDirs = Directory.GetDirectories(mesRootDir)
                .OrderBy(d => Path.GetFileName(d))
                .ToList();

            foreach (var monthDir in monthDirs)
            {
                var monthName = Path.GetFileName(monthDir.TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar));

                // 第三层：该月份下的所有 CSV 文件，按文件名排序（如 25-12-18.csv）
                var csvFiles = Directory.GetFiles(monthDir, "*.csv")
                    .OrderBy(f => Path.GetFileName(f))
                    .ToList();

                foreach (var csvPath in csvFiles)
                {
                    var fileName = Path.GetFileName(csvPath); // 例如 25-12-18.csv

                    // 如果 Redis 中有记录，只处理文件名大于记录值的文件，
                    // 对于同一个文件名，从记录行号之后开始处理
                    if (lastFileName != null)
                    {
                        var cmp = string.Compare(fileName, lastFileName, StringComparison.OrdinalIgnoreCase);
                        if (cmp < 0)
                        {
                            // 该文件名小于已处理文件，跳过
                            Console.WriteLine($"跳过已处理过的旧文件：{fileName}");
                            continue;
                        }
                    }

                    Console.WriteLine($"开始处理文件：{csvPath}");

                    try
                    {
                        int startRow = 0;
                        if (lastFileName != null && string.Equals(fileName, lastFileName, StringComparison.OrdinalIgnoreCase))
                        {
                            // 同一个文件，从上次处理的行号之后开始
                            startRow = lastRow;
                            Console.WriteLine($"检测到上次处理记录：文件 {fileName}，上次处理到第 {lastRow} 行，本次从第 {startRow + 1} 行开始。");
                        }
                        else
                        {
                            Console.WriteLine($"新文件 {fileName}，从第 1 行开始处理。");
                        }

                        // 读取 CSV 文件（从指定行号开始）
                        var (records, actualLastRow) = ReadCsvFile(csvPath, startRow);
                        var recordsList = records.ToList();

                        if (!recordsList.Any())
                        {
                            Console.WriteLine($"文件 {fileName} 没有新数据需要处理。");
                            // 更新 Redis，记录最后查看到的行号，避免重复遍历
                            await redisService.UpdateLastProcessedInfoAsync(fileName, actualLastRow);
                            continue;
                        }

                        // 写入 InfluxDB
                        var inserted = await influxService.WriteRowsAsync(recordsList, database, machineNumber, measurement);
                        Console.WriteLine($"文件 {fileName} 本次新增写入 {inserted} 行。");

                        // 只有当实际写入行数大于0时，才更新 Redis
                        // 如果写入为0（所有数据被过滤），不更新Redis，避免跳过数据
                        if (inserted > 0)
                        {
                            // 更新Redis为文件的总行数，表示这些行已经处理过了
                            // 即使部分数据被过滤（时间条件不满足），也更新Redis
                            // 因为被过滤的数据时间 <= latestTime，可能已经存在于数据库中
                            await redisService.UpdateLastProcessedInfoAsync(fileName, actualLastRow);
                            Console.WriteLine($"已更新 Redis：latest_date(文件名)={fileName}, latest_row={actualLastRow}");
                            
                            if (inserted < recordsList.Count)
                            {
                                Console.WriteLine($"注意：部分数据被过滤（读取 {recordsList.Count} 行，写入 {inserted} 行），这些数据可能已存在或时间条件不满足。");
                            }
                        }
                        else
                        {
                            Console.WriteLine($"所有数据被过滤，未更新 Redis（保持原值：latest_row={lastRow}），下次将继续尝试处理这些数据。");
                        }

                        totalInserted += inserted;
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"处理文件 {csvPath} 时出错: {ex.Message}");
                    }
                }
            }

            return totalInserted;
        }

        /// <summary>
        /// 读取 CSV 文件，支持从指定行号开始读取（跳过前面的行）
        /// </summary>
        /// <param name="filePath">CSV 文件路径</param>
        /// <param name="startRow">开始行号（0 表示从第一行数据开始，跳过表头）</param>
        /// <returns>返回记录集合和最后处理的行号</returns>
        public static (IEnumerable<WeldCsvRecord> records, int lastRow) ReadCsvFile(string filePath, int startRow = 0)
        {
            Console.WriteLine($"Read csv file by CsvHelper (start from row {startRow + 1})...\n");
            var config = new CsvConfiguration(CultureInfo.InvariantCulture)
            {
                HasHeaderRecord = true,
                MissingFieldFound = null
            };

            using var fileStream = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
            using var reader = new StreamReader(fileStream, Encoding.GetEncoding("GB2312"));
            using var csv = new CsvReader(reader, config);

            var allRecords = csv.GetRecords<WeldCsvRecord>().ToList();
            var totalRows = allRecords.Count;

            // 如果 startRow 大于等于总行数，返回空集合
            if (startRow >= totalRows)
            {
                Console.WriteLine($"开始行号 {startRow + 1} 超出文件总行数 {totalRows}，没有新数据。");
                return (Enumerable.Empty<WeldCsvRecord>(), totalRows);
            }

            // 从 startRow 开始读取（跳过前面的行）
            var records = allRecords.Skip(startRow).ToList();
            var lastRow = totalRows - 1; // 最后一行索引（从0开始）

            // 显示前几条记录（最多3条）
            //foreach (var record in records.Take(3))
            //{
            //    Console.WriteLine($"{record.SN}, {record.PointCount}, {record.Endtime}, {record.Temp1}, {record.Temp2}, {record.Temp3}");
            //}

            Console.WriteLine($"共读取 {records.Count} 行数据（从第 {startRow + 1} 行到第 {totalRows} 行）。");

            return (records, lastRow);
        }
    }

    // 记录 CSV 文件格式
    public class WeldCsvRecord
    {
        [Name("条码")]
        public required string SN { get; set; }
        [Name("PointCount")]
        public required string PointCount { get; set; }
        [Name("Endtime")]
        public required string Endtime { get; set; }
        [Name("使用次数")]
        public required string SolderingIron { get; set; }
        [Name("总次数")]
        public required string SolderingJoint { get; set; }
        [Name("焊锡1温度")]
        public required string Temp1 { get; set; }
        [Name("焊锡1时间")]
        public required string Dura1 { get; set; }
        [Name("焊锡1长度")]
        public required string Len1 { get; set; }
        [Name("焊锡1速度")]
        public required string Speed1 { get; set; }
        [Name("焊锡2温度")]
        public required string Temp2 { get; set; }
        [Name("焊锡2时间")]
        public required string Dura2 { get; set; }
        [Name("焊锡2长度")]
        public required string Len2 { get; set; }
        [Name("焊锡2速度")]
        public required string Speed2 { get; set; }
        [Name("焊锡3温度")]
        public required string Temp3 { get; set; }
        [Name("焊锡3时间")]
        public required string Dura3 { get; set; }
        [Name("焊锡3长度")]
        public required string Len3 { get; set; }
        [Name("焊锡3速度")]
        public required string Speed3 { get; set; }
        [Name("焊锡4温度")]
        public required string Temp4 { get; set; }
        [Name("焊锡4时间")]
        public required string Dura4 { get; set; }
        [Name("焊锡4长度")]
        public required string Len4 { get; set; }
        [Name("焊锡4速度")]
        public required string Speed4 { get; set; }
    }
}