using System.Collections.Generic;

namespace ClosedXML.Excel;

/// <summary>
/// Represents the general type of an Excel chart.
/// </summary>
public enum ChartType
{
    Unknown,
    Bar,
    Column,
    Line,
    Pie,
    Area,
    Scatter
}

/// <summary>
/// Represents a single data series within a chart (e.g., a single line in a line chart).
/// </summary>
public class ChartSeries
{
    public string Name { get; internal set; } = string.Empty;
    public ChartType Type { get; internal set; }
    public List<double> Values { get; internal set; } = new();
    public List<string>? Labels { get; internal set; }
    public string? LineColor { get; internal set; }
    public string? FillColor { get; internal set; }
}

/// <summary>
/// A clean, high-level representation of an Excel chart's data, extracted for external use.
/// </summary>
public class Chart
{
    public string Title { get; internal set; } = string.Empty;
    public ChartType Type { get; internal set; }
    public List<string> Labels { get; internal set; } = new(); // For X-Axis / Categories
    public List<ChartSeries> Series { get; internal set; } = new();
}