using System;
using System.Collections.Generic;
using System.IO;
using System.Text.Json;
using System.Threading.Tasks;
using Classin视频解析下载工具.Models;
using Classin视频解析下载工具.Services.CoreServices;
using Classin视频解析下载工具.Services.DownloadServices;
using Classin视频解析下载工具.Shared.Helpers;

namespace Classin视频解析下载工具.Services.ParseServices
{
    public class ParseService : IParseService
    {
        private readonly IValidationService _validationService;
        private readonly IDuplicateDetectionService _duplicateDetectionService;
        private readonly IDownloadService _downloadService;
        private readonly ILoggingService _loggingService;

        public ParseService(
            IValidationService validationService,
            IDuplicateDetectionService duplicateDetectionService,
            IDownloadService downloadService,
            ILoggingService loggingService)
        {
            _validationService = validationService;
            _duplicateDetectionService = duplicateDetectionService;
            _downloadService = downloadService;
            _loggingService = loggingService;
        }

        public async Task<(bool parseSuccess, bool duplicateFound, string? lessonName, string? videoUrl)> ParseContentAsync(string content, List<string> duplicateFiles, List<string> duplicateCourses)
        {
            _loggingService.Debug($"开始解析内容，长度: {content.Length} 字符", "ParseService");

            try
            {
                var result = await TryParseJsonContentAsync(content, duplicateFiles, duplicateCourses);
                _loggingService.Info($"JSON解析{(result.parseSuccess ? "成功" : "失败")}" +
                    (result.duplicateFound ? "，发现重复项" : ""), "ParseService");
                return result;
            }
            catch (Exception ex)
            {
                _loggingService.Warning($"JSON解析失败，回退到行解析: {ex.Message}", "ParseService");
                var result = await UseOriginalLineParsingAsync(content, duplicateFiles, duplicateCourses);
                _loggingService.Info($"行解析{(result.parseSuccess ? "成功" : "失败")}" +
                    (result.duplicateFound ? "，发现重复项" : ""), "ParseService");
                return result;
            }
        }

        private async Task<(bool parseSuccess, bool duplicateFound, string? lessonName, string? videoUrl)> TryParseJsonContentAsync(
            string content, List<string> duplicateFiles, List<string> duplicateCourses)
        {
            _loggingService.Debug("开始尝试JSON解析", "ParseService.JSON");

            try
            {
                using var doc = JsonDocument.Parse(content);
                var root = doc.RootElement;

                if (!root.TryGetProperty("data", out var data))
                {
                    _loggingService.Debug("JSON中未找到data属性，回退到行解析", "ParseService.JSON");
                    return await UseOriginalLineParsingAsync(content, duplicateFiles, duplicateCourses);
                }

                var lessonName = data.TryGetProperty("lessonName", out var lessonNameElement)
                    ? lessonNameElement.GetString() ?? string.Empty
                    : string.Empty;

                if (string.IsNullOrEmpty(lessonName))
                {
                    _loggingService.Debug("未找到有效的课程名称，回退到行解析", "ParseService.JSON");
                    return await UseOriginalLineParsingAsync(content, duplicateFiles, duplicateCourses);
                }

                _loggingService.Debug($"提取到课程名称: {lessonName}", "ParseService.JSON");

                var originalLessonName = lessonName;
                lessonName = FormatUtils.FixEncoding(lessonName);
                if (originalLessonName != lessonName)
                {
                    _loggingService.Debug($"课程名称编码已修复: '{originalLessonName}' -> '{lessonName}'", "ParseService.JSON");
                }

                var videoUrl = ExtractVideoUrlFromJson(data);
                if (string.IsNullOrEmpty(videoUrl))
                {
                    _loggingService.Debug("未找到有效的视频URL，回退到行解析", "ParseService.JSON");
                    return await UseOriginalLineParsingAsync(content, duplicateFiles, duplicateCourses);
                }

                _loggingService.Debug($"提取到视频URL: {videoUrl}", "ParseService.JSON");

                return await ProcessVideoItemAsync(lessonName, videoUrl, duplicateFiles, duplicateCourses);
            }
            catch (Exception ex)
            {
                _loggingService.Warning($"JSON解析过程中发生异常: {ex.Message}，回退到行解析", "ParseService.JSON");
                return await UseOriginalLineParsingAsync(content, duplicateFiles, duplicateCourses);
            }
        }

        private async Task<(bool parseSuccess, bool duplicateFound, string? lessonName, string? videoUrl)> UseOriginalLineParsingAsync(
            string content, List<string> duplicateFiles, List<string> duplicateCourses)
        {
            _loggingService.Debug("开始使用原始行解析方法", "ParseService.Line");

            // 使用更高效的字符串分割方法，减少临时对象创建
            var lines = SplitContentToLines(content);
            _loggingService.Debug($"内容分割为 {lines.Length} 行", "ParseService.Line");

            var currentLessonName = ExtractLessonNameFromLines(lines);

            if (string.IsNullOrEmpty(currentLessonName))
            {
                _loggingService.Warning("未能从行中提取到有效的课程名称", "ParseService.Line");
                return (false, false, null, null);
            }

            _loggingService.Debug($"提取到课程名称: {currentLessonName}", "ParseService.Line");

            var originalLessonName = currentLessonName;
            currentLessonName = FormatUtils.FixEncoding(currentLessonName);
            if (originalLessonName != currentLessonName)
            {
                _loggingService.Debug($"课程名称编码已修复: '{originalLessonName}' -> '{currentLessonName}'", "ParseService.Line");
            }

            var videoUrl = ExtractUrlFromLines(lines);
            if (string.IsNullOrEmpty(videoUrl))
            {
                _loggingService.Warning("未能从行中提取到有效的视频URL", "ParseService.Line");
                return (false, false, null, null);
            }

            _loggingService.Debug($"提取到视频URL: {videoUrl}", "ParseService.Line");

            return await ProcessVideoItemAsync(currentLessonName, videoUrl, duplicateFiles, duplicateCourses);
        }

        /// <summary>
        /// 高效分割内容为行
        /// </summary>
        /// <param name="content">要分割的内容</param>
        /// <returns>分割后的行数组</returns>
        private string[] SplitContentToLines(string content)
        {
            // 使用List<string>来避免预分配过大的数组
            var lines = new System.Collections.Generic.List<string>();
            int startIndex = 0;
            int length = content.Length;

            for (int i = 0; i < length; i++)
            {
                if (content[i] == '\r')
                {
                    // 处理\r\n或单独的\r
                    if (i + 1 < length && content[i + 1] == '\n')
                    {
                        lines.Add(content.Substring(startIndex, i - startIndex));
                        startIndex = i + 2;
                        i++;
                    }
                    else
                    {
                        lines.Add(content.Substring(startIndex, i - startIndex));
                        startIndex = i + 1;
                    }
                }
                else if (content[i] == '\n')
                {
                    // 处理单独的\n
                    lines.Add(content.Substring(startIndex, i - startIndex));
                    startIndex = i + 1;
                }
            }

            // 添加最后一行
            if (startIndex < length)
            {
                lines.Add(content.Substring(startIndex));
            }

            return lines.ToArray();
        }

        private async Task<(bool parseSuccess, bool duplicateFound, string? lessonName, string? videoUrl)> ProcessVideoItemAsync(
            string lessonName, string videoUrl, List<string> duplicateFiles, List<string> duplicateCourses)
        {
            await Task.CompletedTask;

            _loggingService.Debug($"开始处理视频项: {lessonName}", "ParseService.Process");

            if (string.IsNullOrEmpty(lessonName) || string.IsNullOrEmpty(videoUrl))
            {
                _loggingService.Warning($"课程名称或视频URL为空，跳过处理: {lessonName}", "ParseService.Process");
                return (false, false, null, null);
            }

            // 检查重复
            if (_duplicateDetectionService.IsDuplicate(lessonName))
            {
                _loggingService.Info($"检测到重复课程: {lessonName}", "ParseService.Process");
                duplicateCourses.Add(lessonName);
                return (false, true, null, null);
            }

            _loggingService.Info($"视频项处理成功: {lessonName}", "ParseService.Process");
            return (true, false, lessonName, videoUrl);
        }

        public string ExtractLessonNameFromLines(string[] lines)
        {
            _loggingService.Trace("开始从行中提取课程名称", "ParseService.Extract");

            foreach (var line in lines)
            {
                if (line.Contains("lessonName", StringComparison.OrdinalIgnoreCase))
                {
                    var value = ExtractValue(line, "lessonName");
                    if (!string.IsNullOrEmpty(value))
                    {
                        _loggingService.Trace($"成功提取课程名称: {value}", "ParseService.Extract");
                        return value;
                    }
                }
            }

            _loggingService.Trace("未找到课程名称", "ParseService.Extract");
            return string.Empty;
        }

        public string ExtractUrlFromLines(string[] lines)
        {
            _loggingService.Trace("开始从行中提取视频URL", "ParseService.Extract");

            string finalUrl = string.Empty;
            bool playsetEncountered = false;
            bool inFileItem = false;

            foreach (var line in lines)
            {
                if (line.Contains('{') && !inFileItem)
                {
                    inFileItem = true;
                    playsetEncountered = false;
                }
                else if (line.Contains('}') && inFileItem)
                {
                    inFileItem = false;
                }

                if (inFileItem && line.Contains("Playset", StringComparison.OrdinalIgnoreCase))
                {
                    playsetEncountered = true;
                    _loggingService.Trace("检测到Playset字段", "ParseService.Extract");
                }

                if (inFileItem &&
                    line.Contains("url", StringComparison.OrdinalIgnoreCase) &&
                    line.Contains("mp4", StringComparison.OrdinalIgnoreCase))
                {
                    var videoUrl = ExtractValue(line, "url").Replace("\\", "");
                    _loggingService.Trace($"提取到URL候选: {videoUrl}", "ParseService.Extract");

                    if (playsetEncountered)
                    {
                        finalUrl = videoUrl;
                        _loggingService.Trace($"选择最终URL: {finalUrl}", "ParseService.Extract");
                    }
                }
            }

            if (string.IsNullOrEmpty(finalUrl))
            {
                _loggingService.Trace("未找到有效的视频URL", "ParseService.Extract");
            }

            return finalUrl;
        }

        public string ExtractVideoUrlFromJson(JsonElement data)
        {
            _loggingService.Trace("开始从JSON中提取视频URL", "ParseService.Extract");

            if (!data.TryGetProperty("lessonData", out var lessonData))
            {
                _loggingService.Trace("JSON中未找到lessonData属性", "ParseService.Extract");
                return string.Empty;
            }

            if (!lessonData.TryGetProperty("fileList", out var fileList) ||
                fileList.ValueKind != JsonValueKind.Array)
            {
                _loggingService.Trace("JSON中未找到有效的fileList数组", "ParseService.Extract");
                return string.Empty;
            }

            string lastValidUrl = string.Empty;
            int fileCount = 0;

            // 直接使用枚举器，避免创建额外的临时对象
            var fileEnumerator = fileList.EnumerateArray();
            while (fileEnumerator.MoveNext())
            {
                var file = fileEnumerator.Current;
                fileCount++;

                if (file.TryGetProperty("Playset", out var playset) &&
                    playset.ValueKind == JsonValueKind.Array)
                {
                    _loggingService.Trace($"处理第 {fileCount} 个文件项中的Playset数组", "ParseService.Extract");

                    var playEnumerator = playset.EnumerateArray();
                    while (playEnumerator.MoveNext())
                    {
                        var play = playEnumerator.Current;
                        if (play.TryGetProperty("Url", out var urlElement))
                        {
                            var url = urlElement.GetString();
                            if (!string.IsNullOrEmpty(url))
                            {
                                // 只在必要时进行替换操作
                                if (url.Contains('\\'))
                                {
                                    url = url.Replace("\\", "");
                                }

                                if (url.Contains(".mp4", StringComparison.OrdinalIgnoreCase))
                                {
                                    lastValidUrl = url;
                                    _loggingService.Trace($"找到有效MP4 URL: {url}", "ParseService.Extract");
                                }
                            }
                        }
                    }
                }
            }

            if (string.IsNullOrEmpty(lastValidUrl))
            {
                _loggingService.Trace($"在 {fileCount} 个文件项中未找到有效的视频URL", "ParseService.Extract");
            }
            else
            {
                _loggingService.Trace($"成功提取视频URL: {lastValidUrl}", "ParseService.Extract");
            }

            return lastValidUrl;
        }

        private static string ExtractValue(string jsonLine, string key)
        {
            try
            {
                int keyIndex = jsonLine.IndexOf(key, StringComparison.OrdinalIgnoreCase);
                if (keyIndex < 0) return string.Empty;

                int colonIndex = jsonLine.IndexOf(':', keyIndex + key.Length);
                if (colonIndex < 0) return string.Empty;

                int startIndex = colonIndex + 1;
                while (startIndex < jsonLine.Length && char.IsWhiteSpace(jsonLine[startIndex]))
                {
                    startIndex++;
                }

                if (startIndex >= jsonLine.Length) return string.Empty;

                int endIndex;
                char startChar = jsonLine[startIndex];
                if (startChar == '"')
                {
                    // 找到下一个非转义的引号
                    for (endIndex = startIndex + 1; endIndex < jsonLine.Length; endIndex++)
                    {
                        if (jsonLine[endIndex] == '"' && jsonLine[endIndex - 1] != '\\')
                        {
                            break;
                        }
                    }
                    if (endIndex >= jsonLine.Length) endIndex = jsonLine.Length;
                }
                else
                {
                    endIndex = jsonLine.IndexOfAny(new[] { ',', '}', ']' }, startIndex);
                    if (endIndex < 0) endIndex = jsonLine.Length;
                }

                // 提取子字符串并进行最小必要的处理
                string value = jsonLine.Substring(startIndex, endIndex - startIndex);

                // 只在必要时进行Trim操作
                if (char.IsWhiteSpace(value[0]) || char.IsWhiteSpace(value[value.Length - 1]))
                {
                    value = value.Trim();
                }

                // 只在必要时去除引号和逗号
                if (value.Length > 0)
                {
                    if (value[0] == '"' || value[0] == '\'')
                    {
                        value = value.Substring(1);
                    }
                    if (value.Length > 0 && (value[value.Length - 1] == '"' || value[value.Length - 1] == '\'' || value[value.Length - 1] == ','))
                    {
                        value = value.Substring(0, value.Length - 1);
                    }
                }

                return value;
            }
            catch
            {
                return string.Empty;
            }
        }
    }
}