using System;
using System.IO;
using System.Runtime.InteropServices;

namespace weldData_collection
{
    /// <summary>
    /// 网络共享服务类，用于管理网络共享文件夹的连接
    /// </summary>
    public class NetworkShareService
    {
        // Windows API 用于连接网络共享
        [DllImport("mpr.dll", CharSet = CharSet.Unicode, SetLastError = true)]
        private static extern int WNetAddConnection2(
            ref NETRESOURCE netResource,
            string password,
            string username,
            int flags);

        [DllImport("mpr.dll", CharSet = CharSet.Unicode, SetLastError = true)]
        private static extern int WNetCancelConnection2(
            string name,
            int flags,
            bool force);

        [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
        private struct NETRESOURCE
        {
            public int dwScope;
            public int dwType;
            public int dwDisplayType;
            public int dwUsage;
            public string lpLocalName;
            public string lpRemoteName;
            public string lpComment;
            public string lpProvider;
        }

        private const int RESOURCETYPE_DISK = 0x00000001;
        private const int CONNECT_UPDATE_PROFILE = 0x00000001;

        /// <summary>
        /// 尝试检查网络共享是否可访问
        /// </summary>
        /// <param name="networkPath">网络共享路径</param>
        /// <returns>如果可访问返回 true，否则返回 false</returns>
        public static bool TryConnectNetworkShare(string networkPath)
        {
            try
            {
                return Directory.Exists(networkPath);
            }
            catch
            {
                return false;
            }
        }

        /// <summary>
        /// 使用用户名和密码连接网络共享
        /// </summary>
        /// <param name="networkPath">网络共享路径</param>
        /// <param name="username">用户名</param>
        /// <param name="password">密码</param>
        /// <returns>如果连接成功返回 true，否则返回 false</returns>
        public static bool ConnectNetworkShareWithCredentials(string networkPath, string username, string password)
        {
            try
            {
                var netResource = new NETRESOURCE
                {
                    dwType = RESOURCETYPE_DISK,
                    lpRemoteName = networkPath,
                    lpLocalName = null,
                    lpProvider = null
                };

                int result = WNetAddConnection2(ref netResource, password, username, 0);
                
                if (result == 0)
                {
                    Console.WriteLine($"成功使用凭据连接到网络共享: {networkPath}");
                    // 再次验证连接是否成功
                    if (Directory.Exists(networkPath))
                    {
                        return true;
                    }
                }
                else
                {
                    Console.WriteLine($"连接网络共享失败，错误代码: {result}");
                }
                
                return false;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"连接网络共享时发生异常: {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// 确保网络共享已连接，如果直接连接失败则尝试使用凭据连接
        /// </summary>
        /// <param name="networkPath">网络共享路径</param>
        /// <param name="username">用户名（可选）</param>
        /// <param name="password">密码（可选）</param>
        /// <returns>如果连接成功返回 true，否则返回 false</returns>
        public static bool EnsureNetworkShareConnected(string networkPath, string? username = null, string? password = null)
        {
            // 先尝试直接连接
            if (TryConnectNetworkShare(networkPath))
            {
                return true;
            }

            // 如果提供了凭据，尝试使用凭据连接
            if (!string.IsNullOrEmpty(username) && !string.IsNullOrEmpty(password))
            {
                Console.WriteLine($"尝试使用凭据连接网络共享: {networkPath}");
                return ConnectNetworkShareWithCredentials(networkPath, username, password);
            }

            return false;
        }

        /// <summary>
        /// 断开网络共享连接
        /// </summary>
        /// <param name="networkPath">网络共享路径</param>
        /// <param name="force">是否强制断开</param>
        /// <returns>如果断开成功返回 true，否则返回 false</returns>
        public static bool DisconnectNetworkShare(string networkPath, bool force = false)
        {
            try
            {
                int result = WNetCancelConnection2(networkPath, 0, force);
                if (result == 0)
                {
                    Console.WriteLine($"成功断开网络共享连接: {networkPath}");
                    return true;
                }
                else
                {
                    Console.WriteLine($"断开网络共享连接失败，错误代码: {result}");
                    return false;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"断开网络共享连接时发生异常: {ex.Message}");
                return false;
            }
        }
    }
}

