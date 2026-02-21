using System;
using System.Collections.Generic;
using System.IO;
using System.Text.Json;
using System.Threading.Tasks;
using Classin视频解析下载工具.Models;
using Classin视频解析下载工具.Services;

namespace Classin视频解析下载工具.Services
{
    public interface IParseService
    {
        Task<(bool parseSuccess, bool duplicateFound)> ParseContentAsync(string content, List<string> duplicateFiles, List<string> duplicateCourses);
        string ExtractLessonNameFromLines(string[] lines);
        string ExtractUrlFromLines(string[] lines);
        string ExtractVideoUrlFromJson(JsonElement data);
    }

    public class ParseService : IParseService
    {
        private readonly IValidationService _validationService;
        private readonly IDuplicateDetectionService _duplicateDetectionService;
        private readonly IDownloadService _downloadService;

        public ParseService(
            IValidationService validationService,
            IDuplicateDetectionService duplicateDetectionService,
            IDownloadService downloadService)
        {
            _validationService = validationService;
            _duplicateDetectionService = duplicateDetectionService;
            _downloadService = downloadService;
        }

        public async Task<(bool parseSuccess, bool duplicateFound)> ParseContentAsync(string content, List<string> duplicateFiles, List<string> duplicateCourses)
        {
            try
            {
                return await TryParseJsonContentAsync(content, duplicateFiles, duplicateCourses);
            }
            catch
            {
                return await UseOriginalLineParsingAsync(content, duplicateFiles, duplicateCourses);
            }
        }

        private async Task<(bool parseSuccess, bool duplicateFound)> TryParseJsonContentAsync(
            string content, List<string> duplicateFiles, List<string> duplicateCourses)
        {
            try
            {
                using var doc = JsonDocument.Parse(content);
                var root = doc.RootElement;

                if (!root.TryGetProperty("data", out var data))
                {
                    return await UseOriginalLineParsingAsync(content, duplicateFiles, duplicateCourses);
                }

                var lessonName = data.TryGetProperty("lessonName", out var lessonNameElement)
                    ? lessonNameElement.GetString() ?? string.Empty
                    : string.Empty;

                if (string.IsNullOrEmpty(lessonName))
                {
                    return await UseOriginalLineParsingAsync(content, duplicateFiles, duplicateCourses);
                }

                var videoUrl = ExtractVideoUrlFromJson(data);
                if (string.IsNullOrEmpty(videoUrl))
                {
                    return await UseOriginalLineParsingAsync(content, duplicateFiles, duplicateCourses);
                }

                return await ProcessVideoItemAsync(lessonName, videoUrl, duplicateFiles, duplicateCourses);
            }
            catch
            {
                return await UseOriginalLineParsingAsync(content, duplicateFiles, duplicateCourses);
            }
        }

        private async Task<(bool parseSuccess, bool duplicateFound)> UseOriginalLineParsingAsync(
            string content, List<string> duplicateFiles, List<string> duplicateCourses)
        {
            var lines = content.Split(new[] { "\r\n", "\r", "\n" }, StringSplitOptions.None);
            var currentLessonName = ExtractLessonNameFromLines(lines);

            if (string.IsNullOrEmpty(currentLessonName))
            {
                return (false, false);
            }

            var videoUrl = ExtractUrlFromLines(lines);
            if (string.IsNullOrEmpty(videoUrl))
            {
                return (false, false);
            }

            return await ProcessVideoItemAsync(currentLessonName, videoUrl, duplicateFiles, duplicateCourses);
        }

        private async Task<(bool parseSuccess, bool duplicateFound)> ProcessVideoItemAsync(
            string lessonName, string videoUrl, List<string> duplicateFiles, List<string> duplicateCourses)
        {
            if (string.IsNullOrEmpty(lessonName) || string.IsNullOrEmpty(videoUrl))
            {
                return (false, false);
            }

            // 检查重复
            if (_duplicateDetectionService.IsDuplicate(lessonName))
            {
                duplicateCourses.Add(lessonName);
                return (false, true);
            }

            // 添加到重复检测缓存
            _duplicateDetectionService.AddToCache(lessonName);

            // 这里可以添加视频项到下载管理器
            // 由于ParseService不直接引用DownloadManager，
            // 实际的添加逻辑应该在调用方处理

            return (true, false);
        }

        public string ExtractLessonNameFromLines(string[] lines)
        {
            foreach (var line in lines)
            {
                if (line.Contains("lessonName", StringComparison.OrdinalIgnoreCase))
                {
                    return ExtractValue(line, "lessonName");
                }
            }
            return string.Empty;
        }

        public string ExtractUrlFromLines(string[] lines)
        {
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
                }

                if (inFileItem &&
                    line.Contains("url", StringComparison.OrdinalIgnoreCase) &&
                    line.Contains("mp4", StringComparison.OrdinalIgnoreCase))
                {
                    var videoUrl = ExtractValue(line, "url").Replace("\\", "");
                    if (playsetEncountered)
                    {
                        finalUrl = videoUrl;
                    }
                }
            }
            return finalUrl;
        }

        public string ExtractVideoUrlFromJson(JsonElement data)
        {
            if (!data.TryGetProperty("lessonData", out var lessonData)) return string.Empty;
            if (!lessonData.TryGetProperty("fileList", out var fileList) ||
                fileList.ValueKind != JsonValueKind.Array) return string.Empty;

            string lastValidUrl = string.Empty;
            foreach (var file in fileList.EnumerateArray())
            {
                if (file.TryGetProperty("Playset", out var playset) &&
                    playset.ValueKind == JsonValueKind.Array)
                {
                    foreach (var play in playset.EnumerateArray())
                    {
                        if (play.TryGetProperty("Url", out var urlElement))
                        {
                            var url = urlElement.GetString()?.Replace("\\", "") ?? "";
                            if (url.Contains(".mp4", StringComparison.OrdinalIgnoreCase))
                            {
                                lastValidUrl = url;
                            }
                        }
                    }
                }
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
                    endIndex = jsonLine.IndexOf('"', startIndex + 1);
                    if (endIndex < 0) endIndex = jsonLine.Length;
                }
                else
                {
                    endIndex = jsonLine.IndexOfAny(new[] { ',', '}', ']' }, startIndex);
                    if (endIndex < 0) endIndex = jsonLine.Length;
                }

                return jsonLine.Substring(startIndex, endIndex - startIndex)
                    .Trim()
                    .Trim('"', '\'', ',', ' ');
            }
            catch
            {
                return string.Empty;
            }
        }
    }
}