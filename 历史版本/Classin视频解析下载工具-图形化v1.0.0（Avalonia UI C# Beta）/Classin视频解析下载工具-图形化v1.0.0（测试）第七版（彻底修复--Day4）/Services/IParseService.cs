using System;
using System.Collections.Generic;
using System.IO;
using System.Text.Json;
using System.Threading.Tasks;
using Classin视频解析下载工具.Models;

namespace Classin视频解析下载工具.Services
{
    public interface IParseService
    {
        Task<(bool parseSuccess, bool duplicateFound, string? lessonName, string? videoUrl)> ParseContentAsync(string content, List<string> duplicateFiles, List<string> duplicateCourses);
        string ExtractLessonNameFromLines(string[] lines);
        string ExtractUrlFromLines(string[] lines);
        string ExtractVideoUrlFromJson(JsonElement data);
    }
}