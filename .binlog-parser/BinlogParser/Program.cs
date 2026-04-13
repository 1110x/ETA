using System;
using System.Linq;
using Microsoft.Build.Logging.StructuredLogger;

if (args.Length == 0)
{
	Console.WriteLine("Usage: BinlogParser <build.binlog>");
	return;
}

var path = args[0];
var build = BinaryLog.ReadBuild(path);

void Walk(object node, int depth = 0)
{
	var indent = new string(' ', depth * 2);
	var t = node.GetType();
	var nameProp = t.GetProperty("Name") ?? t.GetProperty("Text") ?? t.GetProperty("Header");
	var durationProp = t.GetProperty("Duration") ?? t.GetProperty("DurationMs") ?? t.GetProperty("Elapsed") ?? t.GetProperty("Time");
	var name = nameProp?.GetValue(node)?.ToString() ?? t.Name;
	var dur = durationProp?.GetValue(node)?.ToString() ?? "";
	if (t.Name.IndexOf("Target", StringComparison.OrdinalIgnoreCase) >= 0 ||
		t.Name.IndexOf("Task", StringComparison.OrdinalIgnoreCase) >= 0 ||
		t.Name.IndexOf("Project", StringComparison.OrdinalIgnoreCase) >= 0)
	{
		Console.WriteLine($"{indent}{t.Name}: {name}  Duration={dur}");
	}

	var childrenProp = t.GetProperty("Children");
	if (childrenProp != null)
	{
		var children = childrenProp.GetValue(node) as System.Collections.IEnumerable;
		if (children != null)
		{
			foreach (var child in children)
			{
				Walk(child, depth + 1);
			}
		}
	}
}

Walk(build);
