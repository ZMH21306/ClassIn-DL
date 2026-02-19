using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Classin视频解析下载工具.Models;

namespace Classin视频解析下载工具.Services
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