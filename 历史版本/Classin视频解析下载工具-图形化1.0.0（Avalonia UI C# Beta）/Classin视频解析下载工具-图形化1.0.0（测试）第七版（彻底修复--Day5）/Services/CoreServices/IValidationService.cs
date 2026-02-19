using System;
using System.Collections.Generic;
using System.IO;
using System.Text.RegularExpressions;
using Classin视频解析下载工具.Models;

namespace Classin视频解析下载工具.Services.CoreServices
{
    public class ValidationResult
    {
        public bool IsValid { get; set; }
        public string ErrorMessage { get; set; } = string.Empty;
        public List<string> ValidationErrors { get; set; } = new();
    }

    public interface IValidationService
    {
        ValidationResult ValidateUrl(string url);
        ValidationResult ValidateFilePath(string filePath);
        ValidationResult ValidateDownloadSettings(Models.DownloadSettings settings);
    }
}