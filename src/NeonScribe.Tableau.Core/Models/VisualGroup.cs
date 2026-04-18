namespace NeonScribe.Tableau.Core.Models;

/// <summary>
/// Represents a group of dashboard visuals that share a common title text zone.
/// Titles appear as text zones in the dashboard zone hierarchy and apply to
/// subsequent worksheets until the next title zone is encountered.
/// </summary>
public class VisualGroup
{
    /// <summary>
    /// The title text from the text zone (null if no title zone precedes the worksheets)
    /// </summary>
    public string? Title { get; set; }

    /// <summary>
    /// Visuals (worksheets) belonging to this group, ordered by their position on the dashboard
    /// </summary>
    public List<DashboardVisual> Visuals { get; set; } = new();

    /// <summary>
    /// Position of the group (derived from the title zone, or first visual if no title)
    /// </summary>
    public ZonePosition Position { get; set; } = new();
}
