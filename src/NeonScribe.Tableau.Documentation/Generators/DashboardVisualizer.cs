using System.Text;
using NeonScribe.Tableau.Core.Models;

namespace NeonScribe.Tableau.Documentation.Generators;

/// <summary>
/// Generates SVG visualizations of dashboard layouts
/// </summary>
public class DashboardVisualizer
{
    private const int SVG_SCALE = 100; // Convert Tableau units (hundredths) to pixels
    private const int MAX_DISPLAY_WIDTH = 1000; // Max width for SVG display

    /// <summary>
    /// Generate an SVG visualization of a dashboard layout, wrapped in a container div with a separate legend.
    /// </summary>
    public static string GenerateDashboardSvg(Dashboard dashboard)
    {
        var output = new StringBuilder();

        var displayWidth = Math.Min(dashboard.Size.Width, MAX_DISPLAY_WIDTH);
        var scale = (double)displayWidth / dashboard.Size.Width;
        var displayHeight = (int)(dashboard.Size.Height * scale);

        // --- SVG ---
        var svg = new StringBuilder();
        svg.AppendLine($"<svg width=\"{displayWidth}\" height=\"{displayHeight}\" viewBox=\"0 0 {dashboard.Size.Width} {dashboard.Size.Height}\" xmlns=\"http://www.w3.org/2000/svg\">");

        svg.AppendLine("  <defs>");
        svg.AppendLine("    <style>");
        svg.AppendLine("      .zone-fill { transition: opacity 0.15s; }");
        svg.AppendLine("      .zone-fill:hover { opacity: 0.82; cursor: default; }");
        svg.AppendLine("      .zone-border { fill: none; stroke: rgba(0,0,0,0.18); stroke-width: 1.5; pointer-events: none; }");
        svg.AppendLine("      .layout-border { fill: rgba(0,0,0,0.03); stroke: rgba(0,0,0,0.10); stroke-width: 1; pointer-events: none; }");
        svg.AppendLine("      .zone-type-label { font-family: Arial, sans-serif; font-size: 11px; font-weight: 700; fill: rgba(255,255,255,0.80); dominant-baseline: middle; text-anchor: middle; pointer-events: none; letter-spacing: 0.5px; text-transform: uppercase; }");
        svg.AppendLine("      .zone-name-label { font-family: Arial, sans-serif; font-size: 12px; fill: #ffffff; dominant-baseline: middle; text-anchor: middle; pointer-events: none; }");
        svg.AppendLine("      .zone-type-label-dark { font-family: Arial, sans-serif; font-size: 11px; font-weight: 700; fill: rgba(0,0,0,0.45); dominant-baseline: middle; text-anchor: middle; pointer-events: none; letter-spacing: 0.5px; text-transform: uppercase; }");
        svg.AppendLine("      .zone-name-label-dark { font-family: Arial, sans-serif; font-size: 12px; fill: rgba(0,0,0,0.65); dominant-baseline: middle; text-anchor: middle; pointer-events: none; }");
        svg.AppendLine("      .bg-ws-outline { fill: none; stroke: #f97316; stroke-width: 2; stroke-dasharray: 6,3; pointer-events: none; }");
        svg.AppendLine("      .bg-ws-badge { fill: #f97316; pointer-events: none; }");
        svg.AppendLine("      .bg-ws-label { font-family: Arial, sans-serif; font-size: 10px; font-weight: 700; fill: #ffffff; dominant-baseline: middle; text-anchor: middle; pointer-events: none; letter-spacing: 0.4px; text-transform: uppercase; }");
        svg.AppendLine("      .bg-ws-name { font-family: Arial, sans-serif; font-size: 11px; fill: #f97316; dominant-baseline: middle; text-anchor: middle; pointer-events: none; }");
        svg.AppendLine("      .param-badge { fill: rgba(0,0,0,0.25); pointer-events: none; }");
        svg.AppendLine("      .param-badge-label { font-family: Arial, sans-serif; font-size: 9px; font-weight: 700; fill: #ffffff; dominant-baseline: middle; text-anchor: middle; pointer-events: none; }");
        svg.AppendLine("    </style>");
        svg.AppendLine("  </defs>");

        // Subtle grid background
        svg.AppendLine($"  <rect x=\"0\" y=\"0\" width=\"{dashboard.Size.Width}\" height=\"{dashboard.Size.Height}\" fill=\"#f8f9fa\" rx=\"4\"/>");

        // Layer 1: layout containers (structural guides, behind everything)
        foreach (var zone in dashboard.Zones.Where(z => z.Type.StartsWith("layout", StringComparison.OrdinalIgnoreCase)))
            RenderZone(svg, zone);

        // Layer 2: background worksheets (rendered as outlines, behind content but visible)
        foreach (var zone in dashboard.Zones.Where(z => z.Type.Equals("background-worksheet", StringComparison.OrdinalIgnoreCase)))
            RenderBackgroundWorksheetZone(svg, zone);

        // Layer 3: content zones on top (skip structural/spacer zones)
        foreach (var zone in dashboard.Zones.Where(z => !z.Type.StartsWith("layout", StringComparison.OrdinalIgnoreCase)
                                                      && !z.Type.Equals("background-worksheet", StringComparison.OrdinalIgnoreCase)
                                                      && !z.Type.Equals("skip", StringComparison.OrdinalIgnoreCase)))
            RenderZone(svg, zone);

        svg.AppendLine("</svg>");

        // --- Legend (HTML, outside SVG) ---
        output.AppendLine("<div class=\"dashboard-viz-wrapper\">");
        output.AppendLine("  <div class=\"dashboard-viz-canvas\">");
        output.Append(svg);
        output.AppendLine("  </div>");
        output.AppendLine("  <div class=\"dashboard-viz-legend\">");
        output.AppendLine("    <span class=\"dashboard-viz-legend-title\">Legend</span>");
        output.AppendLine("    <span class=\"dashboard-viz-legend-item\"><span class=\"dashboard-viz-legend-swatch\" style=\"background:#3b82f6\"></span>Worksheet</span>");
        output.AppendLine("    <span class=\"dashboard-viz-legend-item\"><span class=\"dashboard-viz-legend-swatch\" style=\"background:#10b981\"></span>Filter</span>");
        output.AppendLine("    <span class=\"dashboard-viz-legend-item\"><span class=\"dashboard-viz-legend-swatch\" style=\"background:#4b5563\"></span>Title</span>");
        output.AppendLine("    <span class=\"dashboard-viz-legend-item\"><span class=\"dashboard-viz-legend-swatch\" style=\"background:#e5e7eb\"></span>Text</span>");
        output.AppendLine("    <span class=\"dashboard-viz-legend-item\"><span class=\"dashboard-viz-legend-swatch\" style=\"background:#8b5cf6\"></span>Image</span>");
        output.AppendLine("    <span class=\"dashboard-viz-legend-item\"><span class=\"dashboard-viz-legend-swatch\" style=\"background:#b87333\"></span>Legend</span>");
        output.AppendLine("    <span class=\"dashboard-viz-legend-item\"><span class=\"dashboard-viz-legend-swatch\" style=\"background:#f97316\"></span>Background Worksheet</span>");
        output.AppendLine("    <span class=\"dashboard-viz-legend-item\"><span class=\"dashboard-viz-legend-param-badge\">P</span> = Parameter</span>");
        output.AppendLine("  </div>");
        output.AppendLine("</div>");

        return output.ToString();
    }

    private static void RenderZone(StringBuilder svg, DashboardZone zone)
    {
        var x = zone.Position.X / SVG_SCALE;
        var y = zone.Position.Y / SVG_SCALE;
        var width = zone.Position.Width / SVG_SCALE;
        var height = zone.Position.Height / SVG_SCALE;

        if (width < 10 || height < 10)
            return;

        var isLayout = zone.Type.StartsWith("layout", StringComparison.OrdinalIgnoreCase);

        if (isLayout)
        {
            // Layout containers: translucent boundary only, no fill clutter
            svg.AppendLine($"  <rect class=\"layout-border\" x=\"{x}\" y=\"{y}\" width=\"{width}\" height=\"{height}\" rx=\"3\"/>");
            return;
        }

        var fill = GetZoneFill(zone.Type);
        var displayType = GetDisplayType(zone.Type);
        var zoneName = GetZoneName(zone);
        var title = EscapeXml(GetZoneTitle(zone));

        // Dark-text zones (light fills)
        var useDarkText = zone.Type.Equals("text", StringComparison.OrdinalIgnoreCase);
        var typeLabelClass = useDarkText ? "zone-type-label-dark" : "zone-type-label";
        var nameLabelClass = useDarkText ? "zone-name-label-dark" : "zone-name-label";

        var cx = x + width / 2.0;

        svg.AppendLine($"  <g>");
        svg.AppendLine($"    <title>{title}</title>");
        svg.AppendLine($"    <rect class=\"zone-fill\" x=\"{x}\" y=\"{y}\" width=\"{width}\" height=\"{height}\" fill=\"{fill}\" rx=\"4\"/>");
        svg.AppendLine($"    <rect class=\"zone-border\" x=\"{x}\" y=\"{y}\" width=\"{width}\" height=\"{height}\" rx=\"4\"/>");

        // Only add text if the zone is tall enough
        if (height >= 24)
        {
            var maxCharsPerLine = Math.Max(8, (int)(width / 7.5));

            // For short zones use vertical center; for tall zones anchor near the top
            var isShort = height < 60;
            var textY = isShort ? y + height / 2.0 : y + 18;

            if (string.IsNullOrEmpty(zoneName))
            {
                // No name available — fall back to type label only
                svg.AppendLine($"    <text class=\"{typeLabelClass}\" x=\"{cx:F1}\" y=\"{textY:F1}\">{EscapeXml(displayType)}</text>");
            }
            else
            {
                // Show only the name — no type label
                var lines = WrapText(zoneName, maxCharsPerLine);

                if (lines.Count == 1)
                {
                    svg.AppendLine($"    <text class=\"{nameLabelClass}\" x=\"{cx:F1}\" y=\"{textY:F1}\">{EscapeXml(lines[0])}</text>");
                }
                else
                {
                    if (isShort)
                    {
                        // Short zone: two lines centered around midpoint
                        svg.AppendLine($"    <text class=\"{nameLabelClass}\" x=\"{cx:F1}\" y=\"{(y + height / 2.0 - 8):F1}\">{EscapeXml(lines[0])}</text>");
                        svg.AppendLine($"    <text class=\"{nameLabelClass}\" x=\"{cx:F1}\" y=\"{(y + height / 2.0 + 8):F1}\">{EscapeXml(lines[1])}</text>");
                    }
                    else
                    {
                        // Tall zone: stack lines near the top
                        svg.AppendLine($"    <text class=\"{nameLabelClass}\" x=\"{cx:F1}\" y=\"{(y + 18):F1}\">{EscapeXml(lines[0])}</text>");
                        svg.AppendLine($"    <text class=\"{nameLabelClass}\" x=\"{cx:F1}\" y=\"{(y + 34):F1}\">{EscapeXml(lines[1])}</text>");
                    }
                }
            }
        }

        // For parameter-control zones, add a small "P" badge in the top-right corner
        if (zone.Type.Equals("parameter-control", StringComparison.OrdinalIgnoreCase))
        {
            const int pb = 14; // badge size
            var bx = x + width - pb - 3;
            var by = y + 3;
            svg.AppendLine($"    <rect class=\"param-badge\" x=\"{bx}\" y=\"{by}\" width=\"{pb}\" height=\"{pb}\" rx=\"3\"/>");
            svg.AppendLine($"    <text class=\"param-badge-label\" x=\"{bx + pb / 2.0:F1}\" y=\"{by + pb / 2.0:F1}\">P</text>");
        }

        svg.AppendLine($"  </g>");
    }

    /// <summary>
    /// Renders a background worksheet as a dashed orange outline + a small badge,
    /// so it stays visible even when covered by regular content zones.
    /// </summary>
    private static void RenderBackgroundWorksheetZone(StringBuilder svg, DashboardZone zone)
    {
        var x = zone.Position.X / SVG_SCALE;
        var y = zone.Position.Y / SVG_SCALE;
        var width = zone.Position.Width / SVG_SCALE;
        var height = zone.Position.Height / SVG_SCALE;

        if (width < 10 || height < 10)
            return;

        var zoneName = GetZoneName(zone);
        var title = EscapeXml($"Background Worksheet: {zoneName}");

        // Badge: show worksheet name (truncated to fit available width)
        const int badgeH = 18;
        const int badgePadX = 8;
        // Approximate max chars that fit in the badge given zone width, leaving margin
        var maxBadgeChars = Math.Max(4, (int)((width - 16) / 6.5));
        var badgeLabel = string.IsNullOrEmpty(zoneName) ? "BG" : zoneName;
        if (badgeLabel.Length > maxBadgeChars)
            badgeLabel = badgeLabel[..(maxBadgeChars - 1)] + "\u2026";
        // Approximate badge width from label length
        var badgeW = Math.Min((int)(badgeLabel.Length * 6.5) + badgePadX * 2, width - 8);

        svg.AppendLine($"  <g>");
        svg.AppendLine($"    <title>{title}</title>");

        // Dashed outline over the full area
        svg.AppendLine($"    <rect class=\"bg-ws-outline\" x=\"{x + 1}\" y=\"{y + 1}\" width=\"{width - 2}\" height=\"{height - 2}\" rx=\"4\"/>");

        // Badge in the top-left corner with the worksheet name
        var badgeY = y + 4;
        var badgeX = x + 4;
        svg.AppendLine($"    <rect class=\"bg-ws-badge\" x=\"{badgeX}\" y=\"{badgeY}\" width=\"{badgeW}\" height=\"{badgeH}\" rx=\"3\"/>");
        svg.AppendLine($"    <text class=\"bg-ws-label\" x=\"{badgeX + badgeW / 2.0:F1}\" y=\"{badgeY + badgeH / 2.0:F1}\">{EscapeXml(badgeLabel)}</text>");

        svg.AppendLine($"  </g>");
    }

    /// <summary>
    /// Wrap text into at most two lines, each at most maxChars wide. Truncates with ellipsis if needed.
    /// </summary>
    private static List<string> WrapText(string text, int maxChars)
    {
        if (text.Length <= maxChars)
            return [text];

        // Try to break at a space near the midpoint
        var mid = text.Length / 2;
        var breakAt = -1;
        for (int delta = 0; delta <= mid; delta++)
        {
            if (mid - delta > 0 && text[mid - delta] == ' ') { breakAt = mid - delta; break; }
            if (mid + delta < text.Length && text[mid + delta] == ' ') { breakAt = mid + delta; break; }
        }

        string line1, line2;
        if (breakAt > 0)
        {
            line1 = text[..breakAt].Trim();
            line2 = text[(breakAt + 1)..].Trim();
        }
        else
        {
            line1 = text[..maxChars];
            line2 = text[maxChars..];
        }

        // Truncate second line if still too long
        if (line2.Length > maxChars)
            line2 = line2[..(maxChars - 1)] + "\u2026";

        return [line1, line2];
    }

    private static string GetZoneFill(string type)
    {
        return type.ToLower() switch
        {
            "worksheet" => "#3b82f6",
            "parameter-control" => "#10b981",
            "filter" => "#10b981",
            "text" => "#e5e7eb",
            "title" => "#4b5563",
            "bitmap" => "#8b5cf6",
            "background-worksheet" => "#f97316",
            "color-legend" => "#b87333",
            _ => "#9ca3af"
        };
    }

    private static string GetDisplayType(string type)
    {
        return type.ToLower() switch
        {
            "worksheet" => "Worksheet",
            "parameter-control" => "Filter",
            "filter" => "Filter",
            "text" => "Text",
            "title" => "Title",
            "bitmap" => "Image",
            "color-legend" => "Legend",
            "layout-basic" => "Layout",
            "layout-flow" => "Layout",
            _ => type
        };
    }

    private static string GetZoneName(DashboardZone zone)
    {
        if (!string.IsNullOrEmpty(zone.Worksheet))
            return zone.Worksheet;

        if (!string.IsNullOrEmpty(zone.Parameter))
            return zone.Parameter;

        if (!string.IsNullOrEmpty(zone.Content))
        {
            // Strip HTML tags for display
            var text = System.Text.RegularExpressions.Regex.Replace(zone.Content, "<.*?>", string.Empty);
            return text.Trim();
        }

        if (!string.IsNullOrEmpty(zone.ImagePath))
            return System.IO.Path.GetFileName(zone.ImagePath);

        return string.Empty;
    }

    private static string GetZoneTitle(DashboardZone zone)
    {
        var type = GetDisplayType(zone.Type);
        var name = GetZoneName(zone);

        if (!string.IsNullOrEmpty(name))
            return $"{type}: {name}";

        return type;
    }


    private static string EscapeXml(string text)
    {
        if (string.IsNullOrEmpty(text))
            return string.Empty;

        return text
            .Replace("&", "&amp;")
            .Replace("<", "&lt;")
            .Replace(">", "&gt;")
            .Replace("\"", "&quot;")
            .Replace("'", "&apos;");
    }
}
