// =============================================================================
//  ParseExample.cs  ——  C# 中如何解析导出的 JSON（使用 Newtonsoft.Json）
//
//  前置条件：
//   1) 已用本工具生成 Role.cs（与本项目框架的 IConfig / IConfigCategory 等配套）
//   2) 引用 Newtonsoft.Json（NuGet: Install-Package Newtonsoft.Json）
//   3) 把 Json/RoleCategory.json 放到可读取路径（下面示例用相对路径 "Json/RoleCategory.json"）
//
//  编译说明：
//   - 在 Unity 工程里：Role.cs 自带 using UnityEngine; 直接可用，本示例也能编译。
//   - 想做纯 .NET 控制台自测：把 ExcelToJsonCs.py 顶部的 USE_UNITY 改为 False 重新生成
//     Role.cs（去掉 UnityEngine 依赖），再用下面代码即可独立编译运行。
// =============================================================================
using System;
using System.IO;
using Newtonsoft.Json;

class ParseExample
{
    static void Main()
    {
        // 1) 读取并反序列化
        string json = File.ReadAllText("Json/RoleCategory.json");
        RoleCategory cfg = JsonConvert.DeserializeObject<RoleCategory>(json);

        // 2) 构建索引字典（对应 @Role.cs 里的 EndInit：按 Id 建 dict）
        cfg.EndInit();

        // 3) 通过主键 Id 取单条配置
        Role r = cfg.GetOne(5002);
        Console.WriteLine("NickName = " + r.NickName);          // 千里眼

        // 4) 遍历 string[] 数组字段（Excel 中横向多列，导出为真实 JSON 数组）
        Console.WriteLine("Cast 特性:");
        foreach (var c in r.Cast)
            Console.WriteLine("  " + c);                          // 6021 / 6102 / 6062

        // 5) 判存在 / 取全部
        Console.WriteLine("包含 5001? " + cfg.isContain(5001));   // True
        var all = cfg.GetDict();
        Console.WriteLine("共计配置条数: " + all.Count);

        // 6) 遍历整个 dataList
        foreach (var item in cfg.dataList)
            Console.WriteLine($"Id={item.Id} Name={item.NickName}");
    }
}
