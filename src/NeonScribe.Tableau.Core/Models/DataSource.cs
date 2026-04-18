namespace NeonScribe.Tableau.Core.Models;

public class DataSource
{
    public string Name { get; set; } = string.Empty;
    public string InternalName { get; set; } = string.Empty;
    public string Caption { get; set; } = string.Empty;
    public string ConnectionType { get; set; } = string.Empty;
    public string? CustomSql { get; set; }
    public string? ConnectionDetails { get; set; }
    public List<Field> Fields { get; set; } = new();
    public List<AliasMapping> Aliases { get; set; } = new();
}

public class AliasMapping
{
    public string Field { get; set; } = string.Empty;
    public List<AliasEntry> Mappings { get; set; } = new();
}

public class AliasEntry
{
    public string Key { get; set; } = string.Empty;
    public string Value { get; set; } = string.Empty;
}
