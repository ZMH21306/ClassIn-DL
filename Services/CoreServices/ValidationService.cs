using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Classin视频解析下载工具.Models;

namespace Classin视频解析下载工具.Services.CoreServices
{
    public class ValidationService : IValidationService
    {
        public ValidationResult ValidateUrl(string url)
        {
            var result = new ValidationResult();

            if (string.IsNullOrWhiteSpace(url))
            {
                result.IsValid = false;
                result.ErrorMessage = "URL不能为空";
                result.ValidationErrors.Add("URL不能为空");
                return result;
            }

            if (!Uri.TryCreate(url, UriKind.Absolute, out var uri) ||
                (uri.Scheme != Uri.UriSchemeHttp && uri.Scheme != Uri.UriSchemeHttps))
            {
                result.IsValid = false;
                result.ErrorMessage = "无效的URL格式";
                result.ValidationErrors.Add("URL必须是有效的HTTP或HTTPS地址");
                return result;
            }

            // 安全检查：防止 SSRF 攻击，禁止内网地址
            if (IsPrivateOrInternalAddress(uri.Host))
            {
                result.IsValid = false;
                result.ErrorMessage = "禁止访问内网或私有地址";
                result.ValidationErrors.Add($"域名 {uri.Host} 属于内网/私有地址范围");
                return result;
            }

            // 安全检查：域名白名单
            var allowedDomains = new[] { "eeo.cn", "classin.com", "classin.tech" };
            var host = uri.Host.ToLowerInvariant();
            bool isAllowed = allowedDomains.Any(d => host == d || host.EndsWith("." + d, StringComparison.OrdinalIgnoreCase));
            if (!isAllowed)
            {
                result.IsValid = false;
                result.ErrorMessage = "不受信任的域名";
                result.ValidationErrors.Add($"域名 {host} 不在白名单中");
                return result;
            }

            // 安全检查：URL 路径遍历检测
            string urlLower = url.ToLowerInvariant();
            if (urlLower.Contains("../") || urlLower.Contains("..\\") || urlLower.Contains("%2e%2e"))
            {
                result.IsValid = false;
                result.ErrorMessage = "检测到路径遍历攻击";
                result.ValidationErrors.Add("URL包含禁止的路径遍历序列");
                return result;
            }

            result.IsValid = true;
            return result;
        }

        /// <summary>
        /// 检测是否为内网/私有地址（SSRF 防护）
        /// </summary>
        private static bool IsPrivateOrInternalAddress(string host)
        {
            if (string.IsNullOrEmpty(host)) return true;

            // 检查环回地址
            if (host.Equals("localhost", StringComparison.OrdinalIgnoreCase) ||
                host.Equals("127.0.0.1") || host.Equals("::1") ||
                host.StartsWith("127.", StringComparison.OrdinalIgnoreCase))
            {
                return true;
            }

            // 检查私有地址范围
            if (host.StartsWith("10.", StringComparison.OrdinalIgnoreCase) ||
                host.StartsWith("192.168.", StringComparison.OrdinalIgnoreCase) ||
                host.StartsWith("172.16.", StringComparison.OrdinalIgnoreCase) ||
                host.StartsWith("172.17.", StringComparison.OrdinalIgnoreCase) ||
                host.StartsWith("172.18.", StringComparison.OrdinalIgnoreCase) ||
                host.StartsWith("172.19.", StringComparison.OrdinalIgnoreCase) ||
                host.StartsWith("172.2", StringComparison.OrdinalIgnoreCase) ||
                host.StartsWith("172.3", StringComparison.OrdinalIgnoreCase) ||
                host.StartsWith("169.254.", StringComparison.OrdinalIgnoreCase) ||
                host.StartsWith("0.", StringComparison.OrdinalIgnoreCase))
            {
                return true;
            }

            // 检查保留地址范围
            if (host.StartsWith("224.", StringComparison.OrdinalIgnoreCase) ||
                host.StartsWith("239.", StringComparison.OrdinalIgnoreCase) ||
                host.StartsWith("240.", StringComparison.OrdinalIgnoreCase) ||
                host.StartsWith("255.", StringComparison.OrdinalIgnoreCase))
            {
                return true;
            }

            // 检查 IPv6 本地链路/站点本地
            if (host.Contains("::1") || host.Contains("fe80:") ||
                host.Contains("fc00:") || host.Contains("fd00:"))
            {
                return true;
            }

            return false;
        }

        /// <summary>
        /// 验证内容URL（内部服务使用，不限制域名白名单）
        /// </summary>
        public ValidationResult ValidateContentUrl(string url)
        {
            var result = new ValidationResult();

            if (string.IsNullOrWhiteSpace(url))
            {
                result.IsValid = false;
                result.ErrorMessage = "URL不能为空";
                return result;
            }

            if (!Uri.TryCreate(url, UriKind.Absolute, out var uri) ||
                (uri.Scheme != Uri.UriSchemeHttp && uri.Scheme != Uri.UriSchemeHttps))
            {
                result.IsValid = false;
                result.ErrorMessage = "无效的URL格式";
                return result;
            }

            result.IsValid = true;
            return result;
        }

        public ValidationResult ValidateFilePath(string filePath)
        {
            var result = new ValidationResult();

            if (string.IsNullOrWhiteSpace(filePath))
            {
                result.IsValid = false;
                result.ErrorMessage = "文件路径不能为空";
                result.ValidationErrors.Add("文件路径不能为空");
                return result;
            }

            try
            {
                var directory = Path.GetDirectoryName(filePath);
                if (!string.IsNullOrEmpty(directory) && !Directory.Exists(directory))
                {
                    result.IsValid = false;
                    result.ErrorMessage = "目录不存在";
                    result.ValidationErrors.Add($"目录 '{directory}' 不存在");
                    return result;
                }

                var invalidChars = Path.GetInvalidPathChars();
                if (filePath.Any(c => invalidChars.Contains(c)))
                {
                    result.IsValid = false;
                    result.ErrorMessage = "文件路径包含无效字符";
                    result.ValidationErrors.Add("文件路径包含系统不允许的字符");
                    return result;
                }
            }
            catch (Exception ex)
            {
                result.IsValid = false;
                result.ErrorMessage = $"路径验证失败: {ex.Message}";
                result.ValidationErrors.Add($"路径验证异常: {ex.Message}");
                return result;
            }

            result.IsValid = true;
            return result;
        }

        public ValidationResult ValidateDownloadSettings(Models.DownloadSettings settings)
        {
            var result = new ValidationResult();
            var errors = new List<string>();

            if (settings.MaxConcurrentDownloads < 1 || settings.MaxConcurrentDownloads > 100)
            {
                errors.Add("最大并发下载数必须在1-100之间");
            }

            if (string.IsNullOrWhiteSpace(settings.DownloadPath))
            {
                errors.Add("下载路径不能为空");
            }
            else
            {
                var pathValidation = ValidateFilePath(settings.DownloadPath);
                if (!pathValidation.IsValid)
                {
                    errors.AddRange(pathValidation.ValidationErrors);
                }
            }

            if (settings.BufferSizeKB < 1 || settings.BufferSizeKB > 10240)
            {
                errors.Add("缓冲区大小必须在1-10240 KB之间");
            }

            if (settings.TimeoutHours <= 0 || settings.TimeoutHours > 24)
            {
                errors.Add("超时时间必须在0-24小时之间");
            }

            if (settings.MaxRetries < 0 || settings.MaxRetries > 10)
            {
                errors.Add("最大重试次数必须在0-10之间");
            }

            result.IsValid = errors.Count == 0;
            result.ValidationErrors = errors;
            result.ErrorMessage = errors.Count > 0 ? string.Join("; ", errors) : string.Empty;

            return result;
        }
    }
}