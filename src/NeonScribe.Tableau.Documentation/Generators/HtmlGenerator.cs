using System.Text;
using NeonScribe.Tableau.Core.Models;

namespace NeonScribe.Tableau.Documentation.Generators;

public class HtmlGenerator
{
    private readonly TableauWorkbook _workbook;
    private readonly StringBuilder _html;

    public HtmlGenerator(TableauWorkbook workbook)
    {
        _workbook = workbook;
        _html = new StringBuilder();
    }

    public string Generate()
    {
        BuildHtmlDocument();
        return _html.ToString();
    }

    private void BuildHtmlDocument()
    {
        AppendDocumentHeader();
        AppendStyles();
        _html.AppendLine("</head>");
        _html.AppendLine("<body>");

        AppendTableOfContents();
        AppendOverviewSection();
        // Dashboards come first (user-centric view) - these are what users interact with
        AppendDashboardsSection();
        // Technical details follow (collapsed by default in future)
        AppendDataSourcesSection();
        AppendParametersSection();
        AppendWorksheetsSection();
        AppendFieldUsageSection();
        AppendCalculationDependenciesSection();
        AppendJavaScript();

        _html.AppendLine("</body>");
        _html.AppendLine("</html>");
    }

    private void AppendDocumentHeader()
    {
        _html.AppendLine("<!DOCTYPE html>");
        _html.AppendLine("<html lang=\"en\">");
        _html.AppendLine("<head>");
        _html.AppendLine("    <meta charset=\"UTF-8\">");
        _html.AppendLine("    <meta name=\"viewport\" content=\"width=device-width, initial-scale=1.0\">");
        _html.AppendLine($"    <title>Tableau Documentation - {EscapeHtml(_workbook.FileName)}</title>");
        // Premium typography from Google Fonts
        _html.AppendLine("    <link rel=\"preconnect\" href=\"https://fonts.googleapis.com\">");
        _html.AppendLine("    <link rel=\"preconnect\" href=\"https://fonts.gstatic.com\" crossorigin>");
        _html.AppendLine("    <link href=\"https://fonts.googleapis.com/css2?family=Cormorant+Garamond:ital,wght@0,400;0,500;0,600;0,700;1,400&family=DM+Sans:ital,wght@0,400;0,500;0,600;0,700;1,400&family=JetBrains+Mono:wght@400;500&display=swap\" rel=\"stylesheet\">");
    }

    private void AppendStyles()
    {
        _html.AppendLine("    <style>");
        _html.AppendLine(@"
        /* ═══════════════════════════════════════════════════════════════
           LUXURY EDITORIAL REPORT THEME
           A refined, high-end aesthetic for enterprise documentation
           ═══════════════════════════════════════════════════════════════ */

        :root {
            /* Core palette - warm neutrals with copper accent */
            --color-ink: #1a1a1a;
            --color-ink-light: #4a4a4a;
            --color-ink-muted: #737373;
            --color-paper: #fefdfb;
            --color-paper-warm: #faf8f5;
            --color-paper-cool: #f5f4f2;
            --color-border: #e8e6e1;
            --color-border-light: #f0eeea;

            /* Accent colors - sophisticated copper/bronze */
            --color-accent: #b87333;
            --color-accent-light: #d4a574;
            --color-accent-dark: #8b5a2b;
            --color-accent-glow: rgba(184, 115, 51, 0.1);

            /* Semantic colors */
            --color-dimension: #2d5a7b;
            --color-dimension-bg: #e8f1f8;
            --color-measure: #7b5a2d;
            --color-measure-bg: #f8f1e8;
            --color-calculated: #5a2d7b;
            --color-calculated-bg: #f1e8f8;
            --color-filter: #2d7b5a;
            --color-filter-bg: #e8f8f1;
            --color-action: #7b2d5a;
            --color-action-bg: #f8e8f1;
            --color-visual: #3b5a8a;
            --color-visual-bg: #e8eef8;

            /* Typography */
            --font-display: 'Cormorant Garamond', Georgia, serif;
            --font-body: 'DM Sans', -apple-system, sans-serif;
            --font-mono: 'JetBrains Mono', 'Fira Code', monospace;

            /* Spacing scale */
            --space-xs: 0.25rem;
            --space-sm: 0.5rem;
            --space-md: 1rem;
            --space-lg: 1.5rem;
            --space-xl: 2.5rem;
            --space-2xl: 4rem;
            --space-3xl: 6rem;

            /* Transitions */
            --ease-out: cubic-bezier(0.16, 1, 0.3, 1);
            --ease-in-out: cubic-bezier(0.65, 0, 0.35, 1);
        }

        *, *::before, *::after {
            box-sizing: border-box;
            margin: 0;
            padding: 0;
        }

        html {
            font-size: 16px;
            scroll-behavior: smooth;
        }

        body {
            font-family: var(--font-body);
            font-size: 1rem;
            line-height: 1.7;
            color: var(--color-ink);
            background: var(--color-paper);
            -webkit-font-smoothing: antialiased;
            -moz-osx-font-smoothing: grayscale;
        }

        /* Subtle paper texture overlay */
        body::before {
            content: '';
            position: fixed;
            inset: 0;
            background-image: url(""data:image/svg+xml,%3Csvg viewBox='0 0 256 256' xmlns='http://www.w3.org/2000/svg'%3E%3Cfilter id='noise'%3E%3CfeTurbulence type='fractalNoise' baseFrequency='0.8' numOctaves='4' stitchTiles='stitch'/%3E%3C/filter%3E%3Crect width='100%25' height='100%25' filter='url(%23noise)'/%3E%3C/svg%3E"");
            opacity: 0.02;
            pointer-events: none;
            z-index: 1000;
        }

        /* ─────────────────────────────────────────────────────────────────
           CONTAINER & LAYOUT
           ───────────────────────────────────────────────────────────────── */

        .page-wrapper {
            display: flex;
            min-height: 100vh;
        }

        .toc-sidebar {
            position: fixed;
            top: 0;
            left: 0;
            width: 280px;
            height: 100vh;
            background: var(--color-paper-warm);
            border-right: 1px solid var(--color-border);
            overflow-y: auto;
            z-index: 100;
            padding: var(--space-xl) 0;
        }

        /* Decorative accent line on sidebar */
        .toc-sidebar::before {
            content: '';
            position: absolute;
            right: 0;
            top: var(--space-3xl);
            bottom: var(--space-3xl);
            width: 2px;
            background: linear-gradient(
                to bottom,
                transparent,
                var(--color-accent-light) 20%,
                var(--color-accent) 50%,
                var(--color-accent-light) 80%,
                transparent
            );
            opacity: 0.4;
        }

        .container {
            max-width: 1140px;
            margin: 0 auto;
            margin-left: 280px;
            padding: var(--space-2xl) var(--space-xl);
            background: var(--color-paper);
            position: relative;
            flex: 1;
        }

        /* Decorative side border */
        .container::before {
            content: '';
            position: absolute;
            left: 0;
            top: var(--space-3xl);
            bottom: var(--space-3xl);
            width: 1px;
            background: linear-gradient(
                to bottom,
                transparent,
                var(--color-accent-light) 20%,
                var(--color-accent) 50%,
                var(--color-accent-light) 80%,
                transparent
            );
            opacity: 0.4;
        }

        .section {
            /*margin-bottom: var(--space-3xl);
            padding-bottom: var(--space-xl);*/
            animation: fadeInUp 0.6s var(--ease-out) both;
        }

        .section:nth-child(1) { animation-delay: 0.1s; }
        .section:nth-child(2) { animation-delay: 0.15s; }
        .section:nth-child(3) { animation-delay: 0.2s; }
        .section:nth-child(4) { animation-delay: 0.25s; }
        .section:nth-child(5) { animation-delay: 0.3s; }

        @keyframes fadeInUp {
            from {
                opacity: 0;
                transform: translateY(20px);
            }
            to {
                opacity: 1;
                transform: translateY(0);
            }
        }

        /* ─────────────────────────────────────────────────────────────────
           TYPOGRAPHY
           ───────────────────────────────────────────────────────────────── */

        h1, h2, h3, h4 {
            font-family: var(--font-display);
            font-weight: 500;
            letter-spacing: -0.01em;
            color: var(--color-ink);
        }

        h1 {
            font-size: 3.5rem;
            font-weight: 400;
            line-height: 1.1;
            margin-bottom: var(--space-lg);
            position: relative;
            padding-bottom: var(--space-lg);
        }

        h1::after {
            content: '';
            position: absolute;
            bottom: 0;
            left: 0;
            width: 80px;
            height: 3px;
            background: linear-gradient(90deg, var(--color-accent), var(--color-accent-light));
        }

        h2 {
            font-size: 2.25rem;
            margin-top: var(--space-2xl);
            margin-bottom: var(--space-lg);
            padding-bottom: var(--space-md);
            border-bottom: 1px solid var(--color-border);
            position: relative;
        }

        h2::before {
            content: '';
            position: absolute;
            bottom: -1px;
            left: 0;
            width: 40px;
            height: 2px;
            background: var(--color-accent);
        }

        h3 {
            font-size: 1.5rem;
            margin-top: var(--space-xl);
            margin-bottom: var(--space-md);
            color: var(--color-ink);
        }

        h4 {
            font-size: 1.125rem;
            font-family: var(--font-body);
            font-weight: 600;
            margin-top: var(--space-lg);
            margin-bottom: var(--space-sm);
            color: var(--color-ink-light);
            text-transform: uppercase;
            letter-spacing: 0.08em;
            font-size: 0.75rem;
        }

        p {
            margin-bottom: var(--space-md);
            color: var(--color-ink-light);
        }

        strong {
            font-weight: 600;
            color: var(--color-ink);
        }

        /* ─────────────────────────────────────────────────────────────────
           TABLE OF CONTENTS (Fixed Sidebar)
           ───────────────────────────────────────────────────────────────── */

        .toc {
            padding: 0 var(--space-lg);
        }

        .toc h2 {
            font-size: 1.25rem;
            margin: 0 0 var(--space-lg) 0;
            padding: 0 var(--space-sm);
            padding-bottom: var(--space-md);
            border: none;
            border-bottom: 1px solid var(--color-border);
            color: var(--color-ink);
        }

        .toc h2::before {
            display: none;
        }

        .toc ul {
            list-style: none;
            padding: 0;
            margin: 0;
        }

        .toc > ul > li {
            margin-bottom: var(--space-xs);
        }

        .toc ul ul {
            margin-left: 0;
            margin-top: var(--space-xs);
            padding-left: var(--space-md);
            border-left: 2px solid var(--color-border-light);
            margin-left: var(--space-sm);
        }

        .toc a {
            color: var(--color-ink-light);
            text-decoration: none;
            font-size: 0.8125rem;
            transition: all 0.25s var(--ease-out);
            display: block;
            padding: var(--space-xs) var(--space-sm);
            border-radius: 4px;
            position: relative;
            line-height: 1.4;
        }

        .toc a:hover {
            color: var(--color-accent);
            background: var(--color-accent-glow);
        }

        .toc a.active {
            color: var(--color-accent);
            background: var(--color-accent-glow);
            font-weight: 500;
        }

        .toc a.active::before {
            content: '';
            position: absolute;
            left: 0;
            top: 50%;
            transform: translateY(-50%);
            width: 3px;
            height: 60%;
            background: var(--color-accent);
            border-radius: 2px;
        }

        /* TOC Collapsible sections */
        .toc-toggle {
            display: inline-flex;
            align-items: center;
            justify-content: center;
            width: 16px;
            height: 16px;
            margin-right: 4px;
            cursor: pointer;
            color: var(--color-ink-faint);
            transition: transform 0.2s ease, color 0.2s ease;
            flex-shrink: 0;
        }

        .toc-toggle:hover {
            color: var(--color-accent);
        }

        .toc-toggle svg {
            width: 12px;
            height: 12px;
            transition: transform 0.2s ease;
        }

        .toc-item.collapsed > .toc-toggle svg {
            transform: rotate(-90deg);
        }

        .toc-item.collapsed > ul {
            display: none;
        }

        .toc-indented {
            padding-left: var(--space-md);
        }

        .toc-item > .toc-link-wrapper {
            display: flex;
            align-items: center;
        }

        .toc-item > .toc-link-wrapper > a {
            flex: 1;
        }

        /* ─────────────────────────────────────────────────────────────────
           STATISTICS CARDS
           ───────────────────────────────────────────────────────────────── */

        .stats {
            display: grid;
            grid-template-columns: repeat(4, 1fr);
            gap: var(--space-md);
            margin: var(--space-xl) 0;
        }

        @media (max-width: 900px) {
            .stats {
                grid-template-columns: repeat(2, 1fr);
            }
        }

        .stat-card {
            background: var(--color-paper-warm);
            padding: var(--space-lg);
            position: relative;
            overflow: hidden;
            transition: all 0.3s var(--ease-out);
            border: 1px solid var(--color-border-light);
        }

        .stat-card::before {
            content: '';
            position: absolute;
            top: 0;
            left: 0;
            width: 100%;
            height: 3px;
            background: linear-gradient(90deg, var(--color-accent), var(--color-accent-light));
            transform: scaleX(0);
            transform-origin: left;
            transition: transform 0.4s var(--ease-out);
        }

        .stat-card:hover::before {
            transform: scaleX(1);
        }

        .stat-card:hover {
            transform: translateY(-2px);
            box-shadow: 0 8px 24px -8px rgba(0, 0, 0, 0.1);
        }

        .stat-label {
            font-size: 0.6875rem;
            font-weight: 600;
            text-transform: uppercase;
            letter-spacing: 0.12em;
            color: var(--color-ink-muted);
            margin-bottom: var(--space-xs);
        }

        .stat-value {
            font-family: var(--font-display);
            font-size: 2.5rem;
            font-weight: 400;
            color: var(--color-ink);
            line-height: 1;
        }

        .stat-subtitle {
            font-size: 0.75rem;
            color: var(--color-ink-muted);
            margin-top: var(--space-xs);
            font-style: italic;
        }

        .field-annotation {
            display: inline-block;
            font-size: 0.7rem;
            color: var(--color-accent);
            cursor: help;
            margin-left: 0.25rem;
            font-weight: 500;
        }

        .field-annotation:hover {
            color: var(--color-accent-dark, #8b4513);
        }

        /* ─────────────────────────────────────────────────────────────────
           MARK ENCODINGS
           ───────────────────────────────────────────────────────────────── */

        .mark-encodings {
            margin: var(--space-md) 0;
        }

        .mark-encodings h6 {
            font-size: 0.875rem;
            margin-bottom: var(--space-sm);
        }

        .marks-table {
            margin: var(--space-sm) 0;
        }

        .marks-table td {
            padding: var(--space-xs) var(--space-md);
            font-size: 0.875rem;
        }

        .marks-table th {
            padding: var(--space-xs) var(--space-md);
        }

        .kpi-fields-table {
            margin: var(--space-sm) 0;
            width: 100%;
        }

        .kpi-fields-table td {
            padding: var(--space-xs) var(--space-md);
            font-size: 0.875rem;
        }

        .kpi-fields-table th {
            padding: var(--space-xs) var(--space-md);
            text-align: left;
        }

        .kpi-fields-table td:first-child {
            width: 35%;
            font-weight: 500;
        }

        .kpi-fields-table td:nth-child(2) {
            width: 40%;
        }

        .kpi-fields-table td:nth-child(3) {
            width: 25%;
        }

        .kpi-fields-table em {
            color: #666;
            font-style: italic;
        }

        .mark-shelf {
            display: inline-block;
            padding: 2px 8px;
            border-radius: 3px;
            font-size: 0.75rem;
            font-weight: 600;
            text-transform: uppercase;
        }

        .color-shelf {
            background-color: #e8d5d5;
            color: #8b4513;
        }

        .size-shelf {
            background-color: #d5e8d5;
            color: #2e7d32;
        }

        .text-shelf {
            background-color: #d5d5e8;
            color: #3f51b5;
        }

        .shape-shelf {
            background-color: #e8e8d5;
            color: #6d6d00;
        }

        .detail-shelf {
            background-color: #e0e0e0;
            color: #555555;
        }

        .tooltip-shelf {
            background-color: #f5e6d3;
            color: #8b6914;
        }

        /* Map Configuration Styling */
        .map-configuration {
            margin: var(--space-md) 0;
            padding: var(--space-md);
            background-color: #f0f8ff;
            border: 1px solid #b0d4f1;
            border-radius: 4px;
        }

        .map-configuration h4,
        .map-configuration h5,
        .map-configuration h6 {
            color: #1565c0;
            margin-top: var(--space-sm);
            margin-bottom: var(--space-xs);
        }

        .map-configuration h4 {
            font-size: 1rem;
            margin-top: 0;
        }

        .map-configuration h5 {
            font-size: 0.9rem;
        }

        .map-configuration h6 {
            font-size: 0.85rem;
        }

        .geographic-fields-table {
            margin: var(--space-sm) 0;
            width: 100%;
            border-collapse: collapse;
        }

        .geographic-fields-table td,
        .geographic-fields-table th {
            padding: var(--space-xs) var(--space-sm);
            font-size: 0.875rem;
            text-align: left;
        }

        .geographic-fields-table th {
            background-color: #e3f2fd;
            font-weight: 600;
            border-bottom: 2px solid #1565c0;
        }

        .geographic-fields-table tr:nth-child(even) {
            background-color: #f8fcff;
        }

        .geographic-fields-table em {
            color: #666;
            font-style: italic;
            font-size: 0.8rem;
        }

        .map-layers-list {
            list-style: none;
            padding-left: 0;
            margin: var(--space-xs) 0;
        }

        .map-layers-list li {
            padding: var(--space-xs) 0;
            font-size: 0.875rem;
        }

        .map-layers-list li::before {
            content: ""\\2022  "";
            color: #1565c0;
            font-weight: bold;
            margin-right: var(--space-xs);
        }

        .map-configuration p {
            margin: var(--space-xs) 0;
            font-size: 0.875rem;
        }

        /* Table Configuration Styling */
        .table-configuration {
            margin: var(--space-md) 0;
            padding: var(--space-md);
            background-color: #fff9f0;
            border: 1px solid #ffe0b3;
            border-radius: 4px;
        }

        .table-configuration h4,
        .table-configuration h5,
        .table-configuration h6 {
            margin-top: var(--space-sm);
            margin-bottom: var(--space-xs);
            color: #cc6600;
        }

        .table-configuration h4 {
            font-size: 1rem;
            font-weight: 600;
        }

        .table-configuration h6 {
            font-size: 0.875rem;
            font-weight: 600;
            margin-top: var(--space-md);
        }

        .table-configuration p {
            margin: var(--space-xs) 0;
            font-size: 0.8125rem;
        }

        .table-columns-table {
            margin: var(--space-sm) 0;
            width: 100%;
            border-collapse: collapse;
            font-size: 0.8125rem;
        }

        .table-columns-table th,
        .table-columns-table td {
            padding: 6px 8px;
            text-align: left;
            border: 1px solid #ffe0b3;
        }

        .table-columns-table th {
            background-color: #fff3e0;
            font-weight: 600;
            border-bottom: 2px solid #cc6600;
        }

        .table-columns-table tr:nth-child(even) {
            background-color: #fffbf5;
        }

        .table-row-dimensions {
            margin: var(--space-sm) 0;
            padding: var(--space-xs) var(--space-sm);
            background-color: #fffbf5;
            border-left: 3px solid #cc6600;
            font-size: 0.8125rem;
        }

        .table-formatting-grid {
            display: grid;
            grid-template-columns: repeat(2, 1fr);
            gap: var(--space-xs);
            margin: var(--space-sm) 0;
            font-size: 0.8125rem;
        }

        .table-formatting-item {
            padding: 4px 8px;
            background-color: #fffbf5;
            border-left: 2px solid #cc6600;
        }

        .table-formatting-item strong {
            color: #cc6600;
        }

        .color-swatch {
            display: inline-block;
            width: 14px;
            height: 14px;
            border: 1px solid #ccc;
            border-radius: 2px;
            vertical-align: middle;
            margin-left: 4px;
        }

        .visual-description {
            margin: var(--space-sm) 0;
            padding: var(--space-sm) var(--space-md);
            background-color: #f8f9fa;
            border-left: 3px solid var(--accent);
            font-size: 0.9rem;
            line-height: 1.6;
        }

        .visual-description br {
            margin-bottom: 4px;
        }

        .color-swatch {
            display: inline-block;
            width: 14px;
            height: 14px;
            border-radius: 2px;
            border: 1px solid #ccc;
            vertical-align: middle;
            margin-right: 4px;
        }

        .color-theme-info {
            font-size: 0.875rem;
            color: var(--color-ink-muted);
            margin-bottom: var(--space-xs);
        }

        .color-theme-info .color-swatch {
            margin-left: var(--space-sm);
        }

        .color-theme-info .color-swatch:first-of-type {
            margin-left: 0;
        }

        /* ─────────────────────────────────────────────────────────────────
           TABLES
           ───────────────────────────────────────────────────────────────── */

        table {
            width: 100%;
            border-collapse: collapse;
            margin: var(--space-lg) 0;
            font-size: 0.9375rem;
        }

        thead {
            border-bottom: 2px solid var(--color-ink);
        }

        th {
            text-align: left;
            padding: var(--space-sm) var(--space-md);
            font-family: var(--font-body);
            font-size: 0.6875rem;
            font-weight: 600;
            text-transform: uppercase;
            letter-spacing: 0.1em;
            color: var(--color-ink-muted);
        }

        td {
            padding: var(--space-md);
            border-bottom: 1px solid var(--color-border-light);
            vertical-align: top;
            color: var(--color-ink-light);
        }

        tr {
            transition: background-color 0.2s ease;
        }

        tbody tr:hover {
            background-color: var(--color-paper-warm);
        }

        /* ─────────────────────────────────────────────────────────────────
           BADGES
           ───────────────────────────────────────────────────────────────── */

        .badge {
            display: inline-flex;
            align-items: center;
            padding: 0.25rem 0.625rem;
            font-family: var(--font-body);
            font-size: 0.625rem;
            font-weight: 600;
            text-transform: uppercase;
            letter-spacing: 0.08em;
            border-radius: 1px;
            white-space: nowrap;
        }

        .badge-dimension {
            background: var(--color-dimension-bg);
            color: var(--color-dimension);
        }

        .badge-measure {
            background: var(--color-measure-bg);
            color: var(--color-measure);
        }

        .badge-calculated {
            background: var(--color-calculated-bg);
            color: var(--color-calculated);
        }

        .badge-lod {
            background: linear-gradient(135deg, #fce7f3, #f3e8ff);
            color: #7c3aed;
        }

        .badge-table-calc {
            background: var(--color-filter-bg);
            color: var(--color-filter);
        }

        /* ─────────────────────────────────────────────────────────────────
           CONNECTION TYPE
           ───────────────────────────────────────────────────────────────── */

        .connection-type {
            cursor: help;
            border-bottom: 1px dotted var(--color-ink-muted);
        }

        .connection-type-description {
            font-size: 0.8125rem;
            color: var(--color-ink-muted);
            margin-top: var(--space-xs);
            padding-left: var(--space-md);
            border-left: 2px solid var(--color-border);
            font-style: italic;
        }

        /* ─────────────────────────────────────────────────────────────────
           CODE BLOCKS
           ───────────────────────────────────────────────────────────────── */

        .code {
            background: #1e1e1e;
            color: #d4d4d4;
            border-radius: 2px;
            padding: var(--space-lg);
            font-family: var(--font-mono);
            font-size: 0.8125rem;
            line-height: 1.6;
            overflow-x: auto;
            margin: var(--space-md) 0;
            position: relative;
        }

        .code::before {
            content: '';
            position: absolute;
            top: 0;
            left: 0;
            right: 0;
            height: 3px;
            background: linear-gradient(90deg, var(--color-accent), var(--color-accent-light), var(--color-accent));
        }

        .field-ref {
            color: #9cdcfe;
        }

        .calc-link {
            color: var(--color-accent);
            text-decoration: none;
            font-weight: 500;
        }

        .calc-link:hover {
            text-decoration: underline;
            color: var(--color-accent-light);
        }

        .calc-no-link {
            color: var(--color-text-secondary);
            cursor: help;
            border-bottom: 1px dotted var(--color-text-secondary);
        }

        .function {
            color: #dcdcaa;
        }

        .sql-keyword {
            color: #569cd6;
            font-weight: 500;
        }

        .sql-string {
            color: #ce9178;
        }

        .sql-number {
            color: #b5cea8;
        }

        .sql-comment {
            color: #6a9955;
            font-style: italic;
        }

        .sql-identifier {
            color: #4ec9b0;
        }

        .keyword {
            color: #c586c0;
        }

        /* ─────────────────────────────────────────────────────────────────
           CALCULATION EXPLANATION
           ───────────────────────────────────────────────────────────────── */

        .calculation-explanation {
            background: linear-gradient(135deg, var(--color-paper-warm) 0%, var(--color-paper) 100%);
            border-left: 3px solid var(--color-accent);
            padding: var(--space-lg);
            margin: var(--space-lg) 0;
            font-size: 0.9375rem;
            color: var(--color-ink-light);
            position: relative;
        }

        .calculation-explanation::before {
            content: 'INSIGHT';
            position: absolute;
            top: var(--space-sm);
            right: var(--space-md);
            font-size: 0.5625rem;
            font-weight: 700;
            letter-spacing: 0.15em;
            color: var(--color-accent);
            opacity: 0.5;
        }

        .calculation-explanation strong {
            color: var(--color-ink);
        }

        /* ─────────────────────────────────────────────────────────────────
           FILTER & ACTION INFO CARDS
           ───────────────────────────────────────────────────────────────── */

        .filter-info {
            background: var(--color-paper-warm);
            border: 1px solid var(--color-border);
            border-left: 3px solid var(--color-filter);
            padding: var(--space-lg);
            margin: var(--space-md) 0;
            transition: all 0.25s var(--ease-out);
            font-family: var(--font-body);
        }

        .filter-info:hover {
            box-shadow: 0 4px 16px -4px rgba(0, 0, 0, 0.08);
            border-left-color: var(--color-accent);
        }

        .filter-info strong {
            font-family: var(--font-body);
            font-size: 1.125rem;
            font-weight: 600;
            display: block;
            margin-bottom: var(--space-sm);
            color: var(--color-ink);
        }

        .filter-info p {
            font-size: 0.8125rem;
            margin-bottom: var(--space-xs);
            color: var(--color-ink-muted);
        }

        .filter-info p:last-child {
            margin-bottom: 0;
        }

        /* Filters Table View */
        .filters-table {
            width: 100%;
            border-collapse: collapse;
            margin: var(--space-md) 0;
            background: var(--color-paper-warm);
            font-family: var(--font-body);
            font-size: 0.8125rem;
        }

        .filters-table thead {
            background: var(--color-filter);
            color: white;
        }

        .filters-table th {
            padding: var(--space-sm) var(--space-md);
            text-align: left;
            font-weight: 600;
            border: 1px solid var(--color-border);
            color: white;
        }

        .filters-table td {
            padding: var(--space-sm) var(--space-md);
            border: 1px solid var(--color-border);
            color: var(--color-ink-muted);
        }

        .filters-table tbody tr:hover {
            background: #f9f9f9;
        }

        .filters-table a {
            color: var(--color-accent);
            text-decoration: none;
        }

        .filters-table a:hover {
            text-decoration: underline;
        }

        /* Responsive Table Wrapper */
        .table-wrapper {
            overflow-x: auto;
            -webkit-overflow-scrolling: touch;
            margin: var(--space-md) 0;
        }

        .table-wrapper table {
            margin: 0;
        }

        /* ─────────────────────────────────────────────────────────────────
           WORKSHEET ITEMS (within visual groups)
           ───────────────────────────────────────────────────────────────── */

        .worksheet-item {
            background: var(--color-paper-warm);
            border: 1px solid var(--color-border);
            border-left: 3px solid var(--color-visual);
            padding: var(--space-lg);
            margin: var(--space-md) 0;
            transition: all 0.25s var(--ease-out);
        }

        .worksheet-item:hover {
            box-shadow: 0 4px 16px -4px rgba(0, 0, 0, 0.08);
            border-left-color: var(--color-accent);
        }

        .worksheet-item h5 {
            font-family: var(--font-body);
            font-size: 1.125rem;
            font-weight: 600;
            margin: 0 0 var(--space-sm) 0;
            color: var(--color-ink);
        }

        .worksheet-item h6 {
            font-family: var(--font-body);
            font-size: 0.875rem;
            font-weight: 600;
            margin: var(--space-md) 0 var(--space-sm) 0;
            color: var(--color-ink-muted);
        }

        .worksheet-item p {
            font-size: 0.8125rem;
            margin-bottom: var(--space-xs);
            color: var(--color-ink-muted);
        }

        .action-info {
            background: var(--color-paper-warm);
            border: 1px solid var(--color-border);
            border-left: 3px solid var(--color-action);
            padding: var(--space-lg);
            margin: var(--space-md) 0;
            transition: all 0.25s var(--ease-out);
        }

        .action-info:hover {
            box-shadow: 0 4px 16px -4px rgba(0, 0, 0, 0.08);
            border-left-color: var(--color-accent);
        }

        .action-info strong {
            font-family: var(--font-display);
            font-size: 1.125rem;
            font-weight: 500;
            display: block;
            margin-bottom: var(--space-sm);
        }

        .action-info p {
            font-size: 0.875rem;
            margin-bottom: var(--space-xs);
            color: var(--color-ink-muted);
        }

        /* ─────────────────────────────────────────────────────────────────
           SUPPORTING WORKSHEETS TABLE & CLASSIFICATION BADGES
           ───────────────────────────────────────────────────────────────── */

        .supporting-worksheets-table {
            width: 100%;
            margin-bottom: var(--space-lg);
        }

        .supporting-worksheets-table th {
            text-align: left;
            font-size: 0.75rem;
            font-weight: 600;
            color: var(--color-ink-muted);
            text-transform: uppercase;
            letter-spacing: 0.05em;
        }

        .supporting-worksheets-table td {
            padding: var(--space-sm) var(--space-md);
            vertical-align: top;
        }

        .supporting-worksheets-table a {
            color: var(--color-accent);
            text-decoration: none;
            font-weight: 500;
        }

        .supporting-worksheets-table a:hover {
            text-decoration: underline;
        }

        .tooltip-header {
            cursor: help;
            border-bottom: 1px dotted var(--color-ink-muted);
            display: inline-block;
        }

        .worksheet-classification {
            display: inline-block;
            padding: 2px 8px;
            border-radius: 4px;
            font-size: 0.75rem;
            font-weight: 600;
            text-transform: uppercase;
            letter-spacing: 0.03em;
        }

        .worksheet-classification.shared {
            background: #e8f4fd;
            color: #1565c0;
            border: 1px solid #90caf9;
        }

        .worksheet-classification.background {
            background: #f5f5f5;
            color: #616161;
            border: 1px solid #e0e0e0;
            cursor: help;
        }

        /* ─────────────────────────────────────────────────────────────────
           FIELD LINEAGE
           ───────────────────────────────────────────────────────────────── */

        .field-lineage {
            margin-top: var(--space-md);
            padding-top: var(--space-md);
            padding-left: var(--space-md);
            border-top: 1px dashed var(--color-border);
        }

        .field-lineage p {
            font-family: var(--font-body);
            font-size: 0.75rem;
            font-weight: 500;
            color: var(--color-ink-muted);
            margin-bottom: var(--space-xs);
        }

        .lineage-table {
            width: 100%;
            margin: 0;
            font-size: 0.75rem;
            margin-left: var(--space-sm);
        }

        .lineage-table tr:hover {
            background: transparent;
        }

        .lineage-table td {
            padding: 2px var(--space-sm);
            border: none;
            vertical-align: top;
        }

        .lineage-label {
            width: 100px;
            color: var(--color-ink-muted);
            font-weight: 500;
            font-size: 0.6875rem;
        }

        .lineage-value {
            color: var(--color-ink-light);
            font-size: 0.75rem;
        }

        .lineage-value code {
            background: var(--color-paper-cool);
            padding: 0.125rem 0.375rem;
            border-radius: 2px;
            font-family: var(--font-mono);
            font-size: 0.6875rem;
            color: var(--color-ink);
        }

        /* ─────────────────────────────────────────────────────────────────
           DASHBOARD VISUALIZATION
           ───────────────────────────────────────────────────────────────── */

        .dashboard-viz-wrapper {
            margin: var(--space-lg) 0;
            border: 1px solid var(--color-border);
            border-radius: 8px;
            overflow: hidden;
            background: #ffffff;
        }

        .dashboard-viz-canvas {
            background: #f4f5f7;
            padding: var(--space-lg);
            overflow-x: auto;
            display: flex;
            justify-content: center;
        }

        .dashboard-viz-canvas svg {
            max-width: 100%;
            height: auto;
            display: block;
            border-radius: 4px;
            box-shadow: 0 1px 4px rgba(0,0,0,0.10), 0 0 0 1px rgba(0,0,0,0.06);
        }

        .dashboard-viz-legend {
            display: flex;
            flex-wrap: wrap;
            align-items: center;
            gap: 6px 20px;
            padding: 10px var(--space-lg);
            border-top: 1px solid var(--color-border);
            background: #ffffff;
        }

        .dashboard-viz-legend-title {
            font-size: 0.75rem;
            font-weight: 700;
            text-transform: uppercase;
            letter-spacing: 0.06em;
            color: var(--color-ink-muted);
            margin-right: 4px;
        }

        .dashboard-viz-legend-item {
            display: flex;
            align-items: center;
            gap: 6px;
            font-size: 0.8125rem;
            color: var(--color-ink-muted);
            white-space: nowrap;
        }

        .dashboard-viz-legend-swatch {
            display: inline-block;
            width: 12px;
            height: 12px;
            border-radius: 2px;
            flex-shrink: 0;
        }

        .dashboard-viz-legend-param-badge {
            display: inline-flex;
            align-items: center;
            justify-content: center;
            width: 14px;
            height: 14px;
            border-radius: 3px;
            background: rgba(0,0,0,0.25);
            color: #ffffff;
            font-size: 0.625rem;
            font-weight: 700;
            margin: 0 2px;
            vertical-align: middle;
        }

        /* ─────────────────────────────────────────────────────────────────
           USAGE & DEPENDENCY LISTS
           ───────────────────────────────────────────────────────────────── */

        .usage-list {
            list-style: none;
            padding: 0;
            margin: var(--space-sm) 0;
        }

        .usage-list li {
            padding: var(--space-xs) 0;
            padding-left: var(--space-lg);
            position: relative;
            font-size: 0.875rem;
            color: var(--color-ink-light);
        }

        .usage-list li::before {
            content: '';
            position: absolute;
            left: 0;
            top: 50%;
            width: 12px;
            height: 1px;
            background: var(--color-accent);
        }

        .dependency-tree {
            background: var(--color-paper-warm);
            border: 1px solid var(--color-border);
            padding: var(--space-lg);
            margin: var(--space-md) 0;
        }

        .dependency-tree h4 {
            margin-top: 0;
            font-family: var(--font-display);
            font-size: 1.125rem;
            font-weight: 500;
            text-transform: none;
            letter-spacing: normal;
            color: var(--color-ink);
        }

        .dependency-item {
            margin: var(--space-xs) 0;
            padding-left: var(--space-lg);
            position: relative;
            font-size: 0.875rem;
        }

        .dependency-item::before {
            content: '├─';
            position: absolute;
            left: 0;
            color: var(--color-ink-muted);
            font-family: var(--font-mono);
            font-size: 0.75rem;
        }

        /* ─────────────────────────────────────────────────────────────────
           COLLAPSIBLE SECTIONS
           ───────────────────────────────────────────────────────────────── */

        .collapsible-header {
            cursor: pointer;
            user-select: none;
            display: flex;
            align-items: center;
            gap: var(--space-sm);
            padding: var(--space-sm);
            margin: calc(var(--space-sm) * -1);
            margin-bottom: var(--space-sm);
            border-radius: 2px;
            transition: background-color 0.2s ease;
        }

        .collapsible-header:hover {
            background: var(--color-paper-warm);
        }

        .collapsible-header h4 {
            font-family: var(--font-body);
            font-size: 1.125rem;
            font-weight: 600;
            color: var(--color-ink);
        }

        .collapsible-toggle {
            width: 0;
            height: 0;
            border-left: 5px solid transparent;
            border-right: 5px solid transparent;
            border-top: 6px solid var(--color-ink-muted);
            transition: transform 0.3s var(--ease-out);
            flex-shrink: 0;
        }

        .collapsible-toggle.collapsed {
            transform: rotate(-90deg);
        }

        .collapsible-content {
            overflow: hidden;
            transition: max-height 0.4s var(--ease-out), opacity 0.3s ease;
            max-height: 5000px;
            opacity: 1;
        }

        .collapsible-content.collapsed {
            max-height: 0;
            opacity: 0;
        }

        .section-header-row {
            display: flex;
            align-items: center;
            justify-content: space-between;
            margin-bottom: var(--space-md);
        }

        .section-header-row h3 {
            margin: 0;
        }

        .collapse-all-container {
            display: flex;
            justify-content: flex-end;
            margin-bottom: var(--space-md);
        }

        .collapse-all-btn {
            background: transparent;
            border: 1px solid var(--color-border);
            padding: var(--space-xs) var(--space-md);
            font-family: var(--font-body);
            font-size: 0.6875rem;
            font-weight: 600;
            text-transform: uppercase;
            letter-spacing: 0.08em;
            color: var(--color-ink-muted);
            cursor: pointer;
            transition: all 0.25s var(--ease-out);
        }

        .collapse-all-btn:hover {
            background: var(--color-paper-warm);
            border-color: var(--color-accent);
            color: var(--color-accent);
        }

        /* ─────────────────────────────────────────────────────────────────
           EMPTY STATE
           ───────────────────────────────────────────────────────────────── */

        .empty-state {
            text-align: center;
            padding: var(--space-2xl);
            color: var(--color-ink-muted);
            font-style: italic;
        }

        .description {
            color: var(--color-ink-muted);
            font-size: 0.9375rem;
            margin-bottom: var(--space-lg);
        }

        .dashboard-size {
            font-size: 0.875rem;
            color: var(--color-ink-muted);
        }

        /* ─────────────────────────────────────────────────────────────────
           PRINT STYLES
           ───────────────────────────────────────────────────────────────── */

        @media print {
            body {
                background: white;
            }

            body::before {
                display: none;
            }

            .page-wrapper {
                display: block;
            }

            .toc-sidebar {
                position: relative;
                width: 100%;
                height: auto;
                border: 1px solid #ccc;
                margin-bottom: 2rem;
                page-break-after: always;
            }

            .toc-sidebar::before {
                display: none;
            }

            .toc ul ul {
                display: block;
            }

            .container {
                margin-left: 0;
                max-width: 100%;
                padding: 0;
            }

            .container::before {
                display: none;
            }

            .collapsible-content.collapsed {
                max-height: none;
                opacity: 1;
            }

            .collapse-all-container {
                display: none;
            }

            .section {
                page-break-inside: avoid;
            }

            h2, h3 {
                page-break-after: avoid;
            }

            table {
                page-break-inside: avoid;
            }
        }

        /* ─────────────────────────────────────────────────────────────────
           RESPONSIVE ADJUSTMENTS
           ───────────────────────────────────────────────────────────────── */

        @media (max-width: 1024px) {
            .toc-sidebar {
                width: 240px;
            }

            .container {
                margin-left: 240px;
            }
        }

        @media (max-width: 768px) {
            :root {
                --space-xl: 1.5rem;
                --space-2xl: 2.5rem;
                --space-3xl: 3.5rem;
            }

            .page-wrapper {
                flex-direction: column;
            }

            .toc-sidebar {
                position: relative;
                width: 100%;
                height: auto;
                max-height: none;
                border-right: none;
                border-bottom: 1px solid var(--color-border);
                padding: var(--space-lg) 0;
            }

            .toc-sidebar::before {
                display: none;
            }

            .container {
                margin-left: 0;
                max-width: 100%;
            }

            h1 {
                font-size: 2.5rem;
            }

            h2 {
                font-size: 1.75rem;
            }

            .stats {
                grid-template-columns: 1fr;
            }

            .stat-value {
                font-size: 2rem;
            }

            .container::before {
                display: none;
            }

            .toc ul ul {
                display: none;
            }

            .toc > ul {
                display: flex;
                flex-wrap: wrap;
                gap: var(--space-xs);
            }

            .toc > ul > li {
                margin-bottom: 0;
            }

            .toc a {
                font-size: 0.75rem;
                padding: var(--space-xs) var(--space-sm);
                background: var(--color-paper);
                border: 1px solid var(--color-border);
                border-radius: 4px;
            }
        }
        ");
        _html.AppendLine("    </style>");
    }

    private void AppendTableOfContents()
    {
        // Open page wrapper for sidebar + content layout
        _html.AppendLine("<div class=\"page-wrapper\">");

        // Fixed sidebar with TOC
        _html.AppendLine("    <aside class=\"toc-sidebar\">");
        _html.AppendLine("        <nav class=\"toc\">");
        _html.AppendLine("            <h2>Contents</h2>");
        _html.AppendLine("            <ul>");
        _html.AppendLine("                <li><a href=\"#overview\">Overview</a></li>");

        // Dashboards first (user-centric view)
        if (_workbook.Dashboards?.Any() == true)
        {
            _html.AppendLine("                <li class=\"toc-item\">");
            _html.AppendLine("                    <div class=\"toc-link-wrapper\">");
            _html.AppendLine("                        <span class=\"toc-toggle\"><svg viewBox=\"0 0 24 24\" fill=\"none\" stroke=\"currentColor\" stroke-width=\"2\"><polyline points=\"6 9 12 15 18 9\"></polyline></svg></span>");
            _html.AppendLine("                        <a href=\"#dashboards\">Dashboards</a>");
            _html.AppendLine("                    </div>");
            _html.AppendLine("                    <ul>");
            var dashboardIndex = 0;
            foreach (var dashboard in _workbook.Dashboards)
            {
                var dashboardId = ToHtmlId(dashboard.Caption ?? dashboard.Name ?? "dashboard");
                var dashboardName = dashboard.Caption ?? dashboard.Name ?? "Unnamed Dashboard";
                // First dashboard is expanded, subsequent ones are collapsed
                var dashboardCollapsed = dashboardIndex > 0 ? " collapsed" : "";
                _html.AppendLine($"                        <li class=\"toc-item{dashboardCollapsed}\">");
                _html.AppendLine("                            <div class=\"toc-link-wrapper\">");
                _html.AppendLine("                                <span class=\"toc-toggle\"><svg viewBox=\"0 0 24 24\" fill=\"none\" stroke=\"currentColor\" stroke-width=\"2\"><polyline points=\"6 9 12 15 18 9\"></polyline></svg></span>");
                _html.AppendLine($"                                <a href=\"#{dashboardId}\">{EscapeHtml(dashboardName)}</a>");
                _html.AppendLine("                            </div>");
                _html.AppendLine("                            <ul>");

                // Filters sub-section
                if (dashboard.Filters?.Any() == true)
                {
                    _html.AppendLine($"                                <li><a href=\"#{dashboardId}-filters\">Filters</a></li>");
                }

                // Visuals sub-section with grouped visual links
                // First dashboard's Visuals are expanded, subsequent ones are collapsed
                var visualsCollapsed = dashboardIndex > 0 ? " collapsed" : "";
                if (dashboard.VisualGroups?.Any() == true)
                {
                    _html.AppendLine($"                                <li class=\"toc-item{visualsCollapsed}\">");
                    _html.AppendLine("                                    <div class=\"toc-link-wrapper\">");
                    _html.AppendLine("                                        <span class=\"toc-toggle\"><svg viewBox=\"0 0 24 24\" fill=\"none\" stroke=\"currentColor\" stroke-width=\"2\"><polyline points=\"6 9 12 15 18 9\"></polyline></svg></span>");
                    _html.AppendLine($"                                        <a href=\"#{dashboardId}-visuals\">Visuals</a>");
                    _html.AppendLine("                                    </div>");
                    _html.AppendLine("                                    <ul>");
                    foreach (var group in dashboard.VisualGroups)
                    {
                        // Determine the group title for TOC
                        string groupTitle;
                        if (!string.IsNullOrEmpty(group.Title))
                        {
                            groupTitle = group.Title;
                        }
                        else if (group.Visuals.Count == 1)
                        {
                            var singleVisual = group.Visuals[0];
                            groupTitle = singleVisual.Caption ?? singleVisual.Name ?? "Unnamed Visual";
                        }
                        else
                        {
                            groupTitle = "Ungrouped Visuals";
                        }
                        var groupId = $"{dashboardId}-visual-{ToHtmlId(groupTitle)}";
                        _html.AppendLine($"                                        <li><a href=\"#{groupId}\">{EscapeHtml(groupTitle)}</a></li>");
                    }
                    _html.AppendLine("                                    </ul>");
                    _html.AppendLine("                                </li>");
                }
                else if (dashboard.Visuals?.Any() == true)
                {
                    // Fallback: flat visuals if VisualGroups not populated
                    _html.AppendLine($"                                <li class=\"toc-item{visualsCollapsed}\">");
                    _html.AppendLine("                                    <div class=\"toc-link-wrapper\">");
                    _html.AppendLine("                                        <span class=\"toc-toggle\"><svg viewBox=\"0 0 24 24\" fill=\"none\" stroke=\"currentColor\" stroke-width=\"2\"><polyline points=\"6 9 12 15 18 9\"></polyline></svg></span>");
                    _html.AppendLine($"                                        <a href=\"#{dashboardId}-visuals\">Visuals</a>");
                    _html.AppendLine("                                    </div>");
                    _html.AppendLine("                                    <ul>");
                    foreach (var visual in dashboard.Visuals)
                    {
                        var visualName = visual.Caption ?? visual.Name ?? "Unnamed Visual";
                        var visualId = $"{dashboardId}-visual-{ToHtmlId(visualName)}";
                        _html.AppendLine($"                                        <li><a href=\"#{visualId}\">{EscapeHtml(visualName)}</a></li>");
                    }
                    _html.AppendLine("                                    </ul>");
                    _html.AppendLine("                                </li>");
                }

                // Dashboard Layout sub-section
                if (dashboard.Zones?.Any() == true)
                {
                    _html.AppendLine($"                                <li><a href=\"#{dashboardId}-layout\">Dashboard Layout</a></li>");
                }

                _html.AppendLine("                            </ul>");
                _html.AppendLine("                        </li>");
                dashboardIndex++;
            }
            _html.AppendLine("                    </ul>");
            _html.AppendLine("                </li>");
        }

        // Technical details section - expanded by default
        _html.AppendLine("                <li class=\"toc-item\">");
        _html.AppendLine("                    <div class=\"toc-link-wrapper\">");
        _html.AppendLine("                        <span class=\"toc-toggle\"><svg viewBox=\"0 0 24 24\" fill=\"none\" stroke=\"currentColor\" stroke-width=\"2\"><polyline points=\"6 9 12 15 18 9\"></polyline></svg></span>");
        _html.AppendLine("                        <a href=\"#technical-details\">Technical Details</a>");
        _html.AppendLine("                    </div>");
        _html.AppendLine("                    <ul class=\"toc-indented\">");

        if (_workbook.DataSources?.Any() == true)
            _html.AppendLine("                        <li><a href=\"#data-sources\">Data Sources</a></li>");

        if (_workbook.Parameters?.Any() == true)
            _html.AppendLine("                        <li><a href=\"#parameters\">Parameters</a></li>");

        // Check if there are any supporting worksheets (on dashboard but not in Visuals)
        if (HasSupportingWorksheets())
            _html.AppendLine("                        <li><a href=\"#supporting-worksheets\">Supporting Worksheets</a></li>");

        if (_workbook.FieldUsageMap?.Any() == true)
            _html.AppendLine("                        <li><a href=\"#field-usage\">Field Usage</a></li>");

        if (_workbook.CalculationDependencies?.Any() == true)
            _html.AppendLine("                        <li><a href=\"#calculation-dependencies\">Calculation Dependencies</a></li>");

        _html.AppendLine("                    </ul>");
        _html.AppendLine("                </li>");

        _html.AppendLine("            </ul>");
        _html.AppendLine("        </nav>");
        _html.AppendLine("    </aside>");

        // Main content container
        _html.AppendLine("    <div class=\"container\">");
    }

    private string ToHtmlId(string text)
    {
        return System.Text.RegularExpressions.Regex.Replace(
            text.ToLowerInvariant().Replace(" ", "-"),
            @"[^a-z0-9\-]",
            ""
        );
    }

    /// <summary>
    /// Check if there are any supporting worksheets (shared worksheets appearing on 2+ dashboards).
    /// Single-dashboard worksheets without layout-cache are now in Visuals, not Supporting Worksheets.
    /// </summary>
    private bool HasSupportingWorksheets()
    {
        if (_workbook.Dashboards == null || _workbook.Worksheets == null)
            return false;

        // Build count of how many dashboards each worksheet appears on (via zones)
        var allWorksheetNames = _workbook.Worksheets.Select(w => w.Name).ToHashSet();
        var worksheetDashboardCount = new Dictionary<string, int>();

        foreach (var dashboard in _workbook.Dashboards)
        {
            var worksheetsOnThisDashboard = new HashSet<string>();
            if (dashboard.Zones != null)
            {
                foreach (var zone in dashboard.Zones)
                {
                    if (!string.IsNullOrEmpty(zone.Worksheet) && allWorksheetNames.Contains(zone.Worksheet))
                    {
                        worksheetsOnThisDashboard.Add(zone.Worksheet);
                    }
                }
            }

            foreach (var wsName in worksheetsOnThisDashboard)
            {
                if (!worksheetDashboardCount.ContainsKey(wsName))
                    worksheetDashboardCount[wsName] = 0;
                worksheetDashboardCount[wsName]++;
            }
        }

        // Build set of worksheet names that appear in Visuals
        var visualWorksheetNames = new HashSet<string>();
        foreach (var dashboard in _workbook.Dashboards)
        {
            if (dashboard.Visuals != null)
            {
                foreach (var visual in dashboard.Visuals)
                {
                    if (!string.IsNullOrEmpty(visual.WorksheetReference))
                    {
                        visualWorksheetNames.Add(visual.WorksheetReference);
                    }
                }
            }
        }

        // Check if any worksheet appears on 2+ dashboards but not in Visuals (shared worksheets)
        return worksheetDashboardCount.Any(kvp => kvp.Value >= 2 && !visualWorksheetNames.Contains(kvp.Key));
    }

    private void AppendOverviewSection()
    {
        _html.AppendLine("    <section id=\"overview\" class=\"section\">");
        _html.AppendLine($"        <h1>{EscapeHtml(_workbook.FileName)}</h1>");

        _html.AppendLine("        <div class=\"stats\">");
        AppendVersionStatCard();
        AppendStatCard("Worksheets", _workbook.WorksheetsCount.ToString());
        AppendStatCard("Dashboards", _workbook.DashboardsCount.ToString());
        AppendStatCard("Data Sources", _workbook.DataSourcesCount.ToString());
        AppendStatCard("Parameters", _workbook.ParametersCount.ToString());
        AppendStatCard("Total Fields", _workbook.TotalFields.ToString());
        AppendStatCard("Calculated Fields", _workbook.CalculatedFields.ToString());
        AppendStatCard("LOD Calculations", _workbook.LodCalculations.ToString());
        _html.AppendLine("        </div>");

        _html.AppendLine("    </section>");
    }

    private void AppendStatCard(string label, string value)
    {
        _html.AppendLine("            <div class=\"stat-card\">");
        _html.AppendLine($"                <div class=\"stat-label\">{EscapeHtml(label)}</div>");
        _html.AppendLine($"                <div class=\"stat-value\">{EscapeHtml(value)}</div>");
        _html.AppendLine("            </div>");
    }

    private void AppendVersionStatCard()
    {
        var version = _workbook.Version ?? "Unknown";
        var sourceBuild = _workbook.SourceBuild ?? string.Empty;
        var yearName = GetTableauYearName(version, sourceBuild);

        _html.AppendLine("            <div class=\"stat-card\">");
        _html.AppendLine("                <div class=\"stat-label\">Tableau Version</div>");
        // Show year-based name as main value if available, otherwise fall back to version number
        var displayValue = !string.IsNullOrEmpty(yearName) ? yearName : version;
        _html.AppendLine($"                <div class=\"stat-value\">{EscapeHtml(displayValue)}</div>");
        _html.AppendLine("            </div>");
    }

    private string GetTableauYearName(string version, string sourceBuild)
    {
        // Try to extract year-based name from source-build (e.g., "2023.3.0 (20233.23.1017.1841)")
        if (!string.IsNullOrEmpty(sourceBuild))
        {
            var match = System.Text.RegularExpressions.Regex.Match(sourceBuild, @"^(\d{4}\.\d+\.\d+)");
            if (match.Success)
            {
                return $"Tableau {match.Groups[1].Value}";
            }
        }

        // Fall back to mapping version number to approximate year
        // Tableau version numbers map roughly to years:
        // 18.x = 2024.x, 2023.x
        // 10.x = 2017-2018
        if (!string.IsNullOrEmpty(version) && decimal.TryParse(version, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out var versionNum))
        {
            return versionNum switch
            {
                >= 18.0m => "Tableau 2023/2024",
                >= 10.0m and < 18.0m => "Tableau 2017-2022",
                >= 9.0m and < 10.0m => "Tableau 9.x (2015-2016)",
                _ => string.Empty
            };
        }

        return string.Empty;
    }

    private void AppendDataSourcesSection()
    {
        if (_workbook.DataSources == null || !_workbook.DataSources.Any())
            return;

        _html.AppendLine("    <section id=\"data-sources\" class=\"section\">");
        _html.AppendLine("        <h2>Data Sources</h2>");
        _html.AppendLine("        <div class=\"collapse-all-container\">");
        _html.AppendLine("            <button class=\"collapse-all-btn\">Expand All Formulas</button>");
        _html.AppendLine("        </div>");

        foreach (var dataSource in _workbook.DataSources)
        {
            _html.AppendLine($"        <h3>{EscapeHtml(dataSource.Name ?? dataSource.Caption ?? "Unnamed Data Source")}</h3>");

            if (!string.IsNullOrEmpty(dataSource.ConnectionType))
            {
                var connectionTypeDescription = GetConnectionTypeDescription(dataSource.ConnectionType);
                _html.AppendLine($"        <p><strong>Connection Type:</strong> <span class=\"connection-type\" title=\"{EscapeHtml(connectionTypeDescription)}\">{EscapeHtml(dataSource.ConnectionType)}</span></p>");
                if (!string.IsNullOrEmpty(connectionTypeDescription))
                {
                    _html.AppendLine($"        <p class=\"connection-type-description\">{EscapeHtml(connectionTypeDescription)}</p>");
                }
            }

            if (!string.IsNullOrEmpty(dataSource.ConnectionDetails))
            {
                _html.AppendLine($"        <p><strong>Connection:</strong> {EscapeHtml(dataSource.ConnectionDetails)}</p>");
            }

            if (!string.IsNullOrEmpty(dataSource.CustomSql))
            {
                _html.AppendLine("        <h4>Custom SQL</h4>");
                _html.AppendLine($"        <div class=\"code\">{FormatSql(dataSource.CustomSql)}</div>");
            }

            if (dataSource.Fields != null && dataSource.Fields.Any())
            {
                var dsName = dataSource.Name ?? dataSource.Caption ?? "Unnamed";
                _html.AppendLine("        <h4>Fields</h4>");
                _html.AppendLine("        <div class=\"table-wrapper\">");
                _html.AppendLine("        <table>");
                _html.AppendLine("            <thead>");
                _html.AppendLine("                <tr>");
                _html.AppendLine("                    <th>Field Name</th>");
                _html.AppendLine("                    <th>Type</th>");
                _html.AppendLine("                    <th>Role</th>");
                _html.AppendLine("                    <th>Details</th>");
                _html.AppendLine("                </tr>");
                _html.AppendLine("            </thead>");
                _html.AppendLine("            <tbody>");

                foreach (var field in dataSource.Fields)
                {
                    var fieldName = field.Name ?? "Unnamed";
                    var fieldId = $"ds-{ToHtmlId(dsName)}-field-{ToHtmlId(fieldName)}";
                    _html.AppendLine($"                <tr id=\"{fieldId}\">");
                    _html.AppendLine($"                    <td><strong>{EscapeHtml(fieldName)}</strong></td>");
                    _html.AppendLine($"                    <td>{EscapeHtml(field.DataType ?? "unknown")}</td>");

                    // Role badge
                    string roleBadge = field.Role?.ToLower() == "dimension"
                        ? "<span class=\"badge badge-dimension\">Dimension</span>"
                        : "<span class=\"badge badge-measure\">Measure</span>";
                    _html.AppendLine($"                    <td>{roleBadge}</td>");

                    // Details column
                    _html.Append("                    <td>");
                    if (field.IsCalculated)
                    {
                        _html.Append("<span class=\"badge badge-calculated\">Calculated</span>");
                    }
                    _html.AppendLine("</td>");

                    _html.AppendLine("                </tr>");

                    // If calculated, show formula in next row (collapsible)
                    if (field.IsCalculated && !string.IsNullOrEmpty(field.Formula))
                    {
                        _html.AppendLine("                <tr>");
                        _html.AppendLine("                    <td colspan=\"4\">");
                        _html.AppendLine("                        <div class=\"collapsible-header\" style=\"margin:0; padding:4px;\">");
                        _html.AppendLine("                            <span class=\"collapsible-toggle collapsed\"></span>");
                        _html.AppendLine("                            <span><strong>Formula & Details</strong></span>");
                        _html.AppendLine("                        </div>");
                        _html.AppendLine("                        <div class=\"collapsible-content collapsed\">");
                        _html.AppendLine("                            <strong>Formula:</strong>");
                        _html.AppendLine($"                            <div class=\"code\">{FormatFormula(field.Formula)}</div>");

                        // Show explanation for LOD and Table Calculations
                        if (!string.IsNullOrEmpty(field.Explanation))
                        {
                            _html.AppendLine("                            <div class=\"calculation-explanation\">");
                            _html.AppendLine($"                                <strong>📘 What this does:</strong> {EscapeHtml(field.Explanation)}");
                            _html.AppendLine("                            </div>");
                        }

                        if (field.Dependencies != null && field.Dependencies.Any())
                        {
                            _html.AppendLine("                            <strong>Depends on:</strong> ");
                            _html.AppendLine($"                            {string.Join(", ", field.Dependencies.Select(d => $"<span class=\"field-ref\">{EscapeHtml(d)}</span>"))}");
                        }

                        _html.AppendLine("                        </div>");
                        _html.AppendLine("                    </td>");
                        _html.AppendLine("                </tr>");
                    }
                }

                _html.AppendLine("            </tbody>");
                _html.AppendLine("        </table>");
                _html.AppendLine("        </div>");
            }

            // Aliases
            if (dataSource.Aliases != null && dataSource.Aliases.Any())
            {
                _html.AppendLine("        <h4>Field Aliases</h4>");
                foreach (var alias in dataSource.Aliases)
                {
                    _html.AppendLine($"        <p><strong>{EscapeHtml(alias.Field)}:</strong></p>");
                    _html.AppendLine("        <ul>");
                    foreach (var mapping in alias.Mappings)
                    {
                        _html.AppendLine($"            <li>{EscapeHtml(mapping.Key)} → {EscapeHtml(mapping.Value)}</li>");
                    }
                    _html.AppendLine("        </ul>");
                }
            }
        }

        _html.AppendLine("    </section>");
    }

    private void AppendParametersSection()
    {
        if (_workbook.Parameters == null || !_workbook.Parameters.Any())
            return;

        _html.AppendLine("    <section id=\"parameters\" class=\"section\">");
        _html.AppendLine("        <h2>Parameters</h2>");
        _html.AppendLine("        <p class=\"description\">Parameters are user-controllable filters that allow dashboard interactivity.</p>");

        _html.AppendLine("        <div class=\"table-wrapper\">");
        _html.AppendLine("        <table>");
        _html.AppendLine("            <thead>");
        _html.AppendLine("                <tr>");
        _html.AppendLine("                    <th>Parameter Name</th>");
        _html.AppendLine("                    <th>Type</th>");
        _html.AppendLine("                    <th>Current Value</th>");
        _html.AppendLine("                    <th>Allowable Values</th>");
        _html.AppendLine("                </tr>");
        _html.AppendLine("            </thead>");
        _html.AppendLine("            <tbody>");

        foreach (var parameter in _workbook.Parameters)
        {
            // Use Caption first (display name) since that's what filters link to
            var paramName = parameter.Caption ?? parameter.Name ?? "Unnamed";
            var paramId = $"param-{ToHtmlId(paramName)}";
            _html.AppendLine($"                <tr id=\"{paramId}\">");
            _html.AppendLine($"                    <td><strong>{EscapeHtml(paramName)}</strong></td>");
            _html.AppendLine($"                    <td>{EscapeHtml(parameter.DataType ?? "unknown")}</td>");
            _html.AppendLine($"                    <td>{EscapeHtml(parameter.CurrentValue?.ToString() ?? "N/A")}</td>");

            // Allowable values
            _html.Append("                    <td>");
            if (parameter.AllowableValues != null)
            {
                if (parameter.AllowableValues.Type == "range" &&
                    parameter.AllowableValues.Min != null &&
                    parameter.AllowableValues.Max != null)
                {
                    _html.Append($"{parameter.AllowableValues.Min} to {parameter.AllowableValues.Max}");
                }
                else if (parameter.AllowableValues.Type == "list" &&
                         parameter.AllowableValues.Values != null)
                {
                    _html.Append(string.Join(", ", parameter.AllowableValues.Values.Select(v => EscapeHtml(v?.ToString() ?? ""))));
                }
            }
            _html.AppendLine("</td>");

            _html.AppendLine("                </tr>");
        }

        _html.AppendLine("            </tbody>");
        _html.AppendLine("        </table>");
        _html.AppendLine("        </div>");
        _html.AppendLine("    </section>");
    }

    private void AppendWorksheetsSection()
    {
        // Supporting worksheets are shared worksheets that appear on 2+ dashboards.
        // Single-dashboard worksheets without layout-cache are now in Visuals, not here.

        // Build set of worksheet names that appear in Visuals across all dashboards
        var visualWorksheetNames = new HashSet<string>();
        if (_workbook.Dashboards != null)
        {
            foreach (var dashboard in _workbook.Dashboards)
            {
                if (dashboard.Visuals != null)
                {
                    foreach (var visual in dashboard.Visuals)
                    {
                        if (!string.IsNullOrEmpty(visual.WorksheetReference))
                        {
                            visualWorksheetNames.Add(visual.WorksheetReference);
                        }
                    }
                }
            }
        }

        // Build set of worksheet names that appear in dashboard zones (any zone with a name matching a worksheet)
        // Also track which dashboards each worksheet appears on
        var worksheetNamesOnDashboards = new HashSet<string>();
        var worksheetToDashboards = new Dictionary<string, List<string>>();
        if (_workbook.Dashboards != null && _workbook.Worksheets != null)
        {
            var allWorksheetNames = _workbook.Worksheets.Select(w => w.Name).ToHashSet();
            foreach (var dashboard in _workbook.Dashboards)
            {
                if (dashboard.Zones != null)
                {
                    foreach (var zone in dashboard.Zones)
                    {
                        if (!string.IsNullOrEmpty(zone.Worksheet) && allWorksheetNames.Contains(zone.Worksheet))
                        {
                            worksheetNamesOnDashboards.Add(zone.Worksheet);

                            // Track which dashboards use this worksheet
                            if (!worksheetToDashboards.ContainsKey(zone.Worksheet))
                            {
                                worksheetToDashboards[zone.Worksheet] = new List<string>();
                            }
                            var dashboardCaption = dashboard.Caption ?? dashboard.Name ?? "Unknown Dashboard";
                            if (!worksheetToDashboards[zone.Worksheet].Contains(dashboardCaption))
                            {
                                worksheetToDashboards[zone.Worksheet].Add(dashboardCaption);
                            }
                        }
                    }
                }
            }
        }

        // Supporting worksheets = worksheets appearing on 2+ dashboards but not in Visuals (shared worksheets only)
        var supportingWorksheets = _workbook.Worksheets?
            .Where(w => worksheetNamesOnDashboards.Contains(w.Name) &&
                        !visualWorksheetNames.Contains(w.Name) &&
                        worksheetToDashboards.TryGetValue(w.Name, out var dashboards) && dashboards.Count >= 2)
            .ToList();

        if (supportingWorksheets == null || !supportingWorksheets.Any())
            return;

        _html.AppendLine("    <section id=\"supporting-worksheets\" class=\"section\">");
        _html.AppendLine("        <h2>Supporting Worksheets</h2>");
        _html.AppendLine("        <div class=\"collapse-all-container\">");
        _html.AppendLine("            <button class=\"collapse-all-btn\">Collapse All</button>");
        _html.AppendLine("        </div>");

        foreach (var worksheet in supportingWorksheets)
        {
            var worksheetAnchor = $"supporting-worksheet-{ToHtmlId(worksheet.Name ?? "unknown")}";

            // All worksheets in this section are shared (appear on 2+ dashboards)
            var worksheetName = worksheet.Name ?? "unknown";
            var dashboardsUsingWorksheet = worksheetToDashboards.ContainsKey(worksheetName)
                ? worksheetToDashboards[worksheetName]
                : new List<string>();

            _html.AppendLine("        <div class=\"collapsible-header\">");
            _html.AppendLine("            <span class=\"collapsible-toggle\"></span>");
            _html.AppendLine($"            <h3 id=\"{worksheetAnchor}\" style=\"margin:0;\">{EscapeHtml(worksheet.Name ?? worksheet.Caption ?? "Unnamed Worksheet")}</h3>");
            _html.AppendLine("        </div>");
            _html.AppendLine("        <div class=\"collapsible-content\">");

            // Classification badge - always "Shared Worksheet" since we only include 2+ dashboard worksheets
            _html.AppendLine("            <p><span class=\"worksheet-classification shared\">Shared Worksheet</span></p>");
            // Show "Used in:" for shared worksheets
            var dashboardList = string.Join(", ", dashboardsUsingWorksheet.Select(d => EscapeHtml(d)));
            _html.AppendLine($"            <p><strong>Used in:</strong> {dashboardList}</p>");

            // Visual Type - show mark type, with "(Automatic)" suffix if Tableau's mark type was set to Auto
            var displayVisualType = worksheet.VisualType;
            if (displayVisualType == "Automatic" && !string.IsNullOrEmpty(worksheet.MarkType) &&
                !worksheet.MarkType.Equals("Automatic", StringComparison.OrdinalIgnoreCase))
            {
                displayVisualType = $"{worksheet.MarkType} (Automatic)";
            }
            if (!string.IsNullOrEmpty(displayVisualType))
            {
                _html.AppendLine($"            <p><strong>Visual Type:</strong> {EscapeHtml(displayVisualType)}</p>");
            }

            // Customized Label (for KPI/text displays with mixed field references and static text)
            if (worksheet.CustomizedLabel?.FieldRoles?.Any() == true)
            {
                AppendCustomizedLabelForWorksheet(worksheet.CustomizedLabel);
            }

            // Mark Encodings
            AppendMarkEncodingsForWorksheet(worksheet.MarkEncodings);

            // Map Configuration (for map visualizations)
            AppendMapConfigurationForWorksheet(worksheet.MapConfiguration);

            // Table Configuration (for table visualizations)
            AppendTableConfigurationForWorksheet(worksheet.TableConfiguration, worksheet.Title);

            // Fields used
            if (worksheet.FieldsUsed != null && worksheet.FieldsUsed.Any())
            {
                _html.AppendLine("            <h4>Fields Used</h4>");
                _html.AppendLine("            <div class=\"table-wrapper\">");
                _html.AppendLine("            <table>");
                _html.AppendLine("            <thead>");
                _html.AppendLine("                <tr>");
                _html.AppendLine("                    <th>Field</th>");
                _html.AppendLine("                    <th>Shelf</th>");
                _html.AppendLine("                    <th>Aggregation</th>");
                _html.AppendLine("                </tr>");
                _html.AppendLine("            </thead>");
                _html.AppendLine("            <tbody>");

                foreach (var fieldUsage in worksheet.FieldsUsed)
                {
                    _html.AppendLine("                <tr>");
                    _html.AppendLine($"                    <td>{FormatFieldWithAnnotation(fieldUsage.Field ?? "Unknown")}</td>");
                    _html.AppendLine($"                    <td>{EscapeHtml(fieldUsage.Shelf ?? "N/A")}</td>");
                    _html.AppendLine($"                    <td>{EscapeHtml(fieldUsage.Aggregation ?? "None")}</td>");
                    _html.AppendLine("                </tr>");
                }

                _html.AppendLine("            </tbody>");
                _html.AppendLine("            </table>");
                _html.AppendLine("            </div>");
            }

            // Filters applied to this worksheet
            AppendWorksheetFilters(worksheet.Filters);

            // Tooltip
            if (worksheet.Tooltip?.HasCustomTooltip == true && !string.IsNullOrEmpty(worksheet.Tooltip.Content))
            {
                _html.AppendLine("            <h4>Custom Tooltip</h4>");
                // Preserve line breaks in tooltip content
                var tooltipHtml = EscapeHtml(worksheet.Tooltip.Content).Replace("\n", "<br>");
                _html.AppendLine($"            <div class=\"code\">{tooltipHtml}</div>");
            }

            _html.AppendLine("        </div>"); // Close collapsible-content
        }

        _html.AppendLine("    </section>");
    }

    private void AppendDashboardsSection()
    {
        if (_workbook.Dashboards == null || !_workbook.Dashboards.Any())
            return;

        _html.AppendLine("    <section id=\"dashboards\" class=\"section\">");

        foreach (var dashboard in _workbook.Dashboards)
        {
            var dashboardId = ToHtmlId(dashboard.Caption ?? dashboard.Name ?? "dashboard");
            var dashboardTitle = dashboard.Caption ?? dashboard.Name ?? "Unnamed Dashboard";

            // Dashboard title as major section heading
            _html.AppendLine($"        <h2 id=\"{dashboardId}\">{EscapeHtml(dashboardTitle)}</h2>");

            // Show internal name if different from the visible title
            if (!string.IsNullOrEmpty(dashboard.Name) && dashboard.Name != dashboard.Caption)
            {
                _html.AppendLine($"        <p class=\"dashboard-size\"><strong>Internal Name:</strong> {EscapeHtml(dashboard.Name)}</p>");
            }

            if (dashboard.Size != null)
            {
                _html.AppendLine($"        <p class=\"dashboard-size\"><strong>Size:</strong> {dashboard.Size.Width} × {dashboard.Size.Height} pixels</p>");
            }

            // Color Theme section
            if (dashboard.ColorTheme != null)
            {
                AppendDashboardColorTheme(dashboard.ColorTheme);
            }

            // Dashboard Filters section
            if (dashboard.Filters != null && dashboard.Filters.Any())
            {
                _html.AppendLine($"        <h3 id=\"{dashboardId}-filters\">Filters</h3>");

                // Add filters table format
                _html.AppendLine("        <div class=\"table-wrapper\">");
                _html.AppendLine("        <table class=\"filters-table\">");
                _html.AppendLine("            <thead>");
                _html.AppendLine("                <tr>");
                _html.AppendLine("                    <th>Name</th>");
                _html.AppendLine("                    <th>Control</th>");
                _html.AppendLine("                    <th>Default Selection</th>");
                _html.AppendLine("                    <th>Tableau Field</th>");
                _html.AppendLine("                    <th>Base Field</th>");
                _html.AppendLine("                    <th>Data Source</th>");
                _html.AppendLine("                    <th>Field Type</th>");
                _html.AppendLine("                </tr>");
                _html.AppendLine("            </thead>");
                _html.AppendLine("            <tbody>");

                foreach (var filter in dashboard.Filters)
                {
                    _html.AppendLine("                <tr>");

                    // Name
                    _html.AppendLine($"                    <td><strong>{EscapeHtml(filter.Field ?? "Unknown Field")}</strong></td>");

                    // Control
                    var controlType = GetUserFriendlyControlType(filter);
                    _html.AppendLine($"                    <td>{EscapeHtml(controlType ?? "-")}</td>");

                    // Default Selection
                    var defaultSelection = GetDefaultSelection(filter);
                    _html.AppendLine($"                    <td>{EscapeHtml(defaultSelection ?? "-")}</td>");

                    // Tableau Field (from lineage)
                    var tableauField = filter.Lineage?.DisplayName ?? "-";
                    _html.AppendLine($"                    <td>{EscapeHtml(tableauField)}</td>");

                    // Base Field (from lineage, with link if available)
                    var baseField = filter.Lineage?.BaseFieldName ?? "-";
                    if (filter.Lineage != null && !string.IsNullOrEmpty(filter.Lineage.BaseFieldName) &&
                        !string.IsNullOrEmpty(filter.Lineage.DataSourceName))
                    {
                        // Use parameter link if data source is Parameters, otherwise use field link
                        string fieldId;
                        if (filter.Lineage.DataSourceName == "Parameters")
                        {
                            fieldId = $"param-{ToHtmlId(filter.Lineage.BaseFieldName)}";
                        }
                        else
                        {
                            fieldId = $"ds-{ToHtmlId(filter.Lineage.DataSourceName)}-field-{ToHtmlId(filter.Lineage.BaseFieldName)}";
                        }
                        _html.AppendLine($"                    <td><a href=\"#{fieldId}\">{EscapeHtml(baseField)}</a></td>");
                    }
                    else
                    {
                        _html.AppendLine($"                    <td>{EscapeHtml(baseField)}</td>");
                    }

                    // Data Source (with link if it's a parameter, plus allowed values if present)
                    var dataSource = filter.Lineage?.DataSourceName ?? "-";
                    _html.Append("                    <td>");
                    if (dataSource == "Parameters" && !string.IsNullOrEmpty(filter.Lineage?.BaseFieldName))
                    {
                        var paramId = $"param-{ToHtmlId(filter.Lineage.BaseFieldName)}";
                        _html.Append($"<a href=\"#{paramId}\">{EscapeHtml(dataSource)}</a>");
                    }
                    else
                    {
                        _html.Append(EscapeHtml(dataSource));
                    }
                    if (dataSource == "Parameters" && filter.AllowedValues != null && filter.AllowedValues.Any())
                    {
                        _html.Append($"<br/><span style=\"font-size:0.85em;color:#6b7280\">{EscapeHtml(string.Join(", ", filter.AllowedValues))}</span>");
                    }
                    _html.AppendLine("</td>");

                    // Field Type (DataType / Role)
                    var typeInfo = new List<string>();
                    if (!string.IsNullOrEmpty(filter.Lineage?.DataType))
                        typeInfo.Add(filter.Lineage.DataType);
                    if (!string.IsNullOrEmpty(filter.Lineage?.Role))
                        typeInfo.Add(filter.Lineage.Role);
                    var fieldType = typeInfo.Any() ? string.Join(" / ", typeInfo) : "-";
                    _html.AppendLine($"                    <td>{EscapeHtml(fieldType)}</td>");

                    _html.AppendLine("                </tr>");
                }

                _html.AppendLine("            </tbody>");
                _html.AppendLine("        </table>");
                _html.AppendLine("        </div>");
            }

            // Dashboard Visuals section - grouped by title text zones
            if (dashboard.VisualGroups != null && dashboard.VisualGroups.Any())
            {
                _html.AppendLine($"        <div class=\"section-header-row\">");
                _html.AppendLine($"            <h3 id=\"{dashboardId}-visuals\">Visuals</h3>");
                _html.AppendLine("            <button class=\"collapse-all-btn\">Collapse All</button>");
                _html.AppendLine("        </div>");

                foreach (var group in dashboard.VisualGroups)
                {
                    // Determine the group title
                    string groupTitle;
                    if (!string.IsNullOrEmpty(group.Title))
                    {
                        groupTitle = group.Title;
                    }
                    else if (group.Visuals.Count == 1)
                    {
                        // No title zone - use the single visual's caption/name
                        var singleVisual = group.Visuals[0];
                        groupTitle = singleVisual.Caption ?? singleVisual.Name ?? "Unnamed Visual";
                    }
                    else
                    {
                        groupTitle = "Ungrouped Visuals";
                    }

                    var groupId = $"{dashboardId}-visual-{ToHtmlId(groupTitle)}";

                    // Render group header (collapsible)
                    _html.AppendLine("        <div class=\"collapsible-header\">");
                    _html.AppendLine("            <span class=\"collapsible-toggle\"></span>");
                    _html.AppendLine($"            <h4 id=\"{groupId}\" style=\"margin:0;\">{EscapeHtml(groupTitle)}</h4>");
                    _html.AppendLine("        </div>");
                    _html.AppendLine("        <div class=\"collapsible-content\">");

                    // Render each worksheet in the group
                    foreach (var visual in group.Visuals)
                    {
                        var worksheetName = visual.Caption ?? visual.Name ?? "Unnamed Worksheet";
                        var visualId = $"{dashboardId}-worksheet-{ToHtmlId(worksheetName)}";

                        // Show worksheet name as sub-header
                        _html.AppendLine($"            <div class=\"worksheet-item\" id=\"{visualId}\">");
                        _html.AppendLine($"                <h5>{EscapeHtml(worksheetName)}</h5>");

                        // Visual Type - show mark type, with "(Automatic)" suffix if Tableau's mark type was set to Auto
                        var displayVisualType = visual.VisualType;
                        if (displayVisualType == "Automatic" && !string.IsNullOrEmpty(visual.MarkType) &&
                            !visual.MarkType.Equals("Automatic", StringComparison.OrdinalIgnoreCase))
                        {
                            displayVisualType = $"{visual.MarkType} (Automatic)";
                        }
                        if (!string.IsNullOrEmpty(displayVisualType))
                            _html.AppendLine($"                <p><strong>Visual Type:</strong> {EscapeHtml(displayVisualType)}</p>");

                        // For KPI/Big Number visuals with CustomizedLabel, render as a table
                        if (visual.CustomizedLabel?.FieldRoles?.Any() == true &&
                            (displayVisualType == "KPI / Big Number" || displayVisualType == "KPI Metric"))
                        {
                            AppendKpiFieldsTable(visual.CustomizedLabel);
                        }
                        else if (!string.IsNullOrEmpty(visual.Description))
                        {
                            // Multi-line descriptions (for other visuals) use newlines, convert to HTML
                            var descriptionHtml = EscapeHtml(visual.Description).Replace("\n", "<br/>");
                            _html.AppendLine($"                <div class=\"visual-description\">{descriptionHtml}</div>");
                        }

                        // Mark Encodings section
                        AppendMarkEncodings(visual.MarkEncodings, visual.HasActualTooltipContent);

                        // Map Configuration (for map visualizations)
                        AppendMapConfiguration(visual.MapConfiguration);

                        // Table Configuration (for table visualizations)
                        AppendTableConfiguration(visual.TableConfiguration, visual.Title);

                        // Detailed Fields Used table (from worksheet)
                        if (visual.DetailedFieldsUsed != null && visual.DetailedFieldsUsed.Any())
                        {
                            _html.AppendLine("                <h6>Fields Used</h6>");
                            _html.AppendLine("                <div class=\"table-wrapper\">");
                            _html.AppendLine("                <table>");
                            _html.AppendLine("                <thead>");
                            _html.AppendLine("                    <tr>");
                            _html.AppendLine("                        <th>Field</th>");
                            _html.AppendLine("                        <th>Shelf</th>");
                            _html.AppendLine("                        <th>Aggregation</th>");
                            _html.AppendLine("                    </tr>");
                            _html.AppendLine("                </thead>");
                            _html.AppendLine("                <tbody>");

                            foreach (var fieldUsage in visual.DetailedFieldsUsed)
                            {
                                _html.AppendLine("                    <tr>");
                                _html.AppendLine($"                        <td>{FormatFieldWithAnnotation(fieldUsage.Field ?? "Unknown")}</td>");
                                _html.AppendLine($"                        <td>{EscapeHtml(fieldUsage.Shelf ?? "N/A")}</td>");
                                _html.AppendLine($"                        <td>{EscapeHtml(fieldUsage.Aggregation ?? "None")}</td>");
                                _html.AppendLine("                    </tr>");
                            }

                            _html.AppendLine("                </tbody>");
                            _html.AppendLine("                </table>");
                            _html.AppendLine("                </div>");
                        }
                        else if (visual.FieldsUsed != null && visual.FieldsUsed.Any())
                        {
                            // Fallback to simple list if detailed info not available
                            _html.AppendLine("                <p><strong>Fields Used:</strong></p>");
                            _html.AppendLine("                <ul class=\"usage-list\">");
                            foreach (var field in visual.FieldsUsed)
                            {
                                _html.AppendLine($"                    <li>{EscapeHtml(field)}</li>");
                            }
                            _html.AppendLine("                </ul>");
                        }

                        // Custom Tooltip
                        if (visual.Tooltip?.HasCustomTooltip == true && !string.IsNullOrEmpty(visual.Tooltip.Content))
                        {
                            _html.AppendLine("                <h6>Custom Tooltip</h6>");
                            var tooltipHtml = EscapeHtml(visual.Tooltip.Content).Replace("\n", "<br>");
                            _html.AppendLine($"                <div class=\"code\">{tooltipHtml}</div>");
                        }

                        // Visual-specific filters (filters that only apply to this visual)
                        if (visual.VisualSpecificFilters != null && visual.VisualSpecificFilters.Any())
                        {
                            _html.AppendLine("                <p><strong>Visual-Specific Filters:</strong></p>");
                            foreach (var filter in visual.VisualSpecificFilters)
                            {
                                _html.AppendLine("                <div class=\"filter-info\" style=\"margin-left: 20px;\">");
                                _html.AppendLine($"                    <strong>{EscapeHtml(filter.Field ?? "Unknown Field")}</strong>");

                                // Control line first (user-facing), then Type
                                var controlType = GetUserFriendlyControlType(filter);
                                if (!string.IsNullOrEmpty(controlType))
                                    _html.AppendLine($"                    <p>Control: {EscapeHtml(controlType)}</p>");

                                _html.AppendLine($"                    <p>Type: {EscapeHtml(filter.FilterType ?? filter.Type ?? "Unknown")}</p>");

                                // Default selection - show "All" for multi-select filters without specific default
                                var defaultSelection = GetDefaultSelection(filter);
                                if (!string.IsNullOrEmpty(defaultSelection))
                                    _html.AppendLine($"                    <p>Default Selection: {EscapeHtml(defaultSelection)}</p>");

                                if (!string.IsNullOrEmpty(filter.Notes))
                                    _html.AppendLine($"                    <p><em>Notes: {EscapeHtml(filter.Notes)}</em></p>");

                                // Display field lineage with links
                                AppendFieldLineage(filter.Lineage, "                    ", filter.AllowedValues);

                                _html.AppendLine("                </div>");
                            }
                        }

                        if (visual.IsSharedAcrossDashboards)
                            _html.AppendLine("                <p><em>This visual is shared across multiple dashboards</em></p>");

                        _html.AppendLine("            </div>");
                    }

                    _html.AppendLine("        </div>");
                }
            }
            else if (dashboard.Visuals != null && dashboard.Visuals.Any())
            {
                // Fallback: render flat visuals if VisualGroups not populated (backward compatibility)
                _html.AppendLine($"        <h3 id=\"{dashboardId}-visuals\">Visuals</h3>");
                foreach (var visual in dashboard.Visuals)
                {
                    var visualTitle = visual.Caption ?? visual.Name ?? "Unnamed Visual";
                    _html.AppendLine($"        <p>{EscapeHtml(visualTitle)}</p>");
                }
            }

            // Supporting Worksheets subsection (only shared worksheets that appear on 2+ dashboards)
            if (dashboard.SupportingWorksheets != null && dashboard.SupportingWorksheets.Any())
            {
                _html.AppendLine($"        <h3 id=\"{dashboardId}-supporting-worksheets\" class=\"tooltip-header\" title=\"Worksheets shared across multiple dashboards. These appear on 2 or more dashboards.\">Supporting Worksheets</h3>");
                _html.AppendLine("        <div class=\"table-wrapper\">");
                _html.AppendLine("        <table class=\"supporting-worksheets-table\">");
                _html.AppendLine("            <thead>");
                _html.AppendLine("                <tr>");
                _html.AppendLine("                    <th>NAME</th>");
                _html.AppendLine("                    <th>VISUAL TYPE</th>");
                _html.AppendLine("                    <th>USED IN</th>");
                _html.AppendLine("                </tr>");
                _html.AppendLine("            </thead>");
                _html.AppendLine("            <tbody>");

                foreach (var supportingWs in dashboard.SupportingWorksheets)
                {
                    var worksheetAnchor = $"supporting-worksheet-{ToHtmlId(supportingWs.Name)}";
                    var displayName = supportingWs.Caption ?? supportingWs.Name;
                    var usedInList = supportingWs.UsedInDashboards?.Any() == true
                        ? string.Join(", ", supportingWs.UsedInDashboards.Select(d => EscapeHtml(d)))
                        : "—";

                    _html.AppendLine("                <tr>");
                    _html.AppendLine($"                    <td><a href=\"#{worksheetAnchor}\">{EscapeHtml(displayName)}</a></td>");
                    _html.AppendLine($"                    <td>{EscapeHtml(supportingWs.VisualType)}</td>");
                    _html.AppendLine($"                    <td>{usedInList}</td>");
                    _html.AppendLine("                </tr>");
                }

                _html.AppendLine("            </tbody>");
                _html.AppendLine("        </table>");
                _html.AppendLine("        </div>");
            }

            // Dashboard Layout Visualization
            if (dashboard.Zones != null && dashboard.Zones.Any())
            {
                _html.AppendLine($"        <h3 id=\"{dashboardId}-layout\">Dashboard Layout</h3>");
                _html.AppendLine(DashboardVisualizer.GenerateDashboardSvg(dashboard));
            }

            // Actions
            if (dashboard.Actions != null && dashboard.Actions.Any())
            {
                _html.AppendLine("        <h3>Actions</h3>");
                foreach (var action in dashboard.Actions)
                {
                    _html.AppendLine("        <div class=\"action-info\">");
                    _html.AppendLine($"            <strong>{EscapeHtml(action.Caption ?? action.Name ?? "Unnamed Action")}</strong>");
                    _html.AppendLine($"            <p>Type: {EscapeHtml(action.Type ?? "Unknown")}</p>");
                    _html.AppendLine($"            <p>Trigger: {EscapeHtml(action.Trigger ?? "Unknown")}</p>");

                    if (action.Source != null && !string.IsNullOrEmpty(action.Source.Name))
                        _html.AppendLine($"            <p>Source: {EscapeHtml(action.Source.Name)}</p>");

                    if (action.Target != null && !string.IsNullOrEmpty(action.Target.Name))
                        _html.AppendLine($"            <p>Target: {EscapeHtml(action.Target.Name)}</p>");

                    if (action.Fields != null && action.Fields.Any())
                        _html.AppendLine($"            <p>Fields: {string.Join(", ", action.Fields.Select(f => EscapeHtml(f)))}</p>");

                    _html.AppendLine("        </div>");
                }
            }
        }

        _html.AppendLine("    </section>");

        // Add technical details section marker
        _html.AppendLine("    <section id=\"technical-details\" class=\"section\">");
        _html.AppendLine("        <h2>Technical Details</h2>");
        _html.AppendLine("        <p class=\"description\">The following sections contain technical information about data sources, worksheets, and field usage.</p>");
        _html.AppendLine("    </section>");
    }

    private void AppendFieldUsageSection()
    {
        if (_workbook.FieldUsageMap == null || !_workbook.FieldUsageMap.Any())
            return;

        _html.AppendLine("    <section id=\"field-usage\" class=\"section\">");
        _html.AppendLine("        <h2>Field Usage Cross-Reference</h2>");
        _html.AppendLine("        <p class=\"description\">This section shows where each field is used throughout the workbook.</p>");

        _html.AppendLine("        <div class=\"table-wrapper\">");
        _html.AppendLine("        <table>");
        _html.AppendLine("            <thead>");
        _html.AppendLine("                <tr>");
        _html.AppendLine("                    <th>Field Name</th>");
        _html.AppendLine("                    <th>Data Source</th>");
        _html.AppendLine("                    <th>Usage</th>");
        _html.AppendLine("                </tr>");
        _html.AppendLine("            </thead>");
        _html.AppendLine("            <tbody>");

        foreach (var usage in _workbook.FieldUsageMap)
        {
            _html.AppendLine("                <tr>");
            _html.AppendLine($"                    <td><strong>{EscapeHtml(usage.Key)}</strong></td>");
            _html.AppendLine($"                    <td>{EscapeHtml(usage.Value.DataSource ?? "Unknown")}</td>");

            // Usage column
            _html.Append("                    <td>");
            var usages = new List<string>();

            if (usage.Value.UsedInWorksheets?.Any() == true)
                usages.Add($"Worksheets: {string.Join(", ", usage.Value.UsedInWorksheets)}");

            if (usage.Value.UsedInFilters?.Any() == true)
                usages.Add($"Filters: {string.Join(", ", usage.Value.UsedInFilters)}");

            if (usage.Value.UsedInCalculations?.Any() == true)
                usages.Add($"Calculations: {string.Join(", ", usage.Value.UsedInCalculations)}");

            if (usage.Value.UsedInActions?.Any() == true)
                usages.Add($"Actions: {string.Join(", ", usage.Value.UsedInActions)}");

            if (usages.Any())
                _html.Append(EscapeHtml(string.Join(" | ", usages)));
            else
                _html.Append("<em>Not used</em>");

            _html.AppendLine("</td>");
            _html.AppendLine("                </tr>");
        }

        _html.AppendLine("            </tbody>");
        _html.AppendLine("        </table>");
        _html.AppendLine("        </div>");
        _html.AppendLine("    </section>");
    }

    private void AppendCalculationDependenciesSection()
    {
        if (_workbook.CalculationDependencies == null || !_workbook.CalculationDependencies.Any())
            return;

        _html.AppendLine("    <section id=\"calculation-dependencies\" class=\"section\">");
        _html.AppendLine("        <h2>Calculation Dependencies</h2>");
        _html.AppendLine("        <p class=\"description\">This section shows how calculated fields depend on other fields.</p>");

        foreach (var calc in _workbook.CalculationDependencies)
        {
            _html.AppendLine("        <div class=\"dependency-tree\">");
            _html.AppendLine($"            <h4>{EscapeHtml(calc.Key)}</h4>");

            if (!string.IsNullOrEmpty(calc.Value.Formula))
            {
                _html.AppendLine("            <strong>Formula:</strong>");
                _html.AppendLine($"            <div class=\"code\">{FormatFormula(calc.Value.Formula)}</div>");
            }

            // Show explanation for LOD and Table Calculations
            if (!string.IsNullOrEmpty(calc.Value.Explanation))
            {
                _html.AppendLine("            <div class=\"calculation-explanation\">");
                _html.AppendLine($"                <strong>📘 What this does:</strong> {EscapeHtml(calc.Value.Explanation)}");
                _html.AppendLine("            </div>");
            }

            if (calc.Value.Dependencies?.Any() == true)
            {
                _html.AppendLine("            <strong>Depends on:</strong>");
                _html.AppendLine("            <ul class=\"usage-list\">");
                foreach (var dep in calc.Value.Dependencies)
                {
                    _html.AppendLine($"                <li>{EscapeHtml(dep)}</li>");
                }
                _html.AppendLine("            </ul>");
            }

            if (calc.Value.UsedBy?.Any() == true)
            {
                _html.AppendLine("            <strong>Used by:</strong>");
                _html.AppendLine("            <ul class=\"usage-list\">");
                foreach (var usedBy in calc.Value.UsedBy)
                {
                    _html.AppendLine($"                <li>{EscapeHtml(usedBy)}</li>");
                }
                _html.AppendLine("            </ul>");
            }

            _html.AppendLine("        </div>");
        }

        _html.AppendLine("    </section>");
        _html.AppendLine("    </div>"); // Close container
        _html.AppendLine("</div>"); // Close page-wrapper
    }

    private string FormatFormula(string formula)
    {
        if (string.IsNullOrEmpty(formula))
            return string.Empty;

        // Escape HTML first
        string formatted = EscapeHtml(formula);

        // Highlight field references [FieldName]
        formatted = System.Text.RegularExpressions.Regex.Replace(
            formatted,
            @"\[([^\]]+)\]",
            "<span class=\"field-ref\">[$1]</span>"
        );

        // Highlight common functions
        var functions = new[] { "SUM", "AVG", "COUNT", "MIN", "MAX", "IF", "THEN", "ELSE", "END",
                                "FIXED", "INCLUDE", "EXCLUDE", "RUNNING_SUM", "WINDOW_AVG" };
        foreach (var func in functions)
        {
            formatted = System.Text.RegularExpressions.Regex.Replace(
                formatted,
                $@"\b{func}\b",
                $"<span class=\"function\">{func}</span>",
                System.Text.RegularExpressions.RegexOptions.IgnoreCase
            );
        }

        return formatted;
    }

    private string FormatSql(string sql)
    {
        if (string.IsNullOrEmpty(sql))
            return string.Empty;

        // Don't escape HTML first - work with original string to handle quotes properly
        string formatted = sql;

        // Highlight string literals (single quotes) BEFORE escaping HTML
        // Use a placeholder to protect them
        var stringLiterals = new System.Collections.Generic.List<string>();
        formatted = System.Text.RegularExpressions.Regex.Replace(
            formatted,
            @"'([^']*)'",
            match => {
                stringLiterals.Add(match.Value);
                return $"___STRING_LITERAL_{stringLiterals.Count - 1}___";
            }
        );

        // Highlight comments (-- style) BEFORE escaping HTML
        var comments = new System.Collections.Generic.List<string>();
        formatted = System.Text.RegularExpressions.Regex.Replace(
            formatted,
            @"(--[^\n]*)",
            match => {
                comments.Add(match.Value);
                return $"___COMMENT_{comments.Count - 1}___";
            }
        );

        // Now escape HTML
        formatted = EscapeHtml(formatted);

        // Highlight SQL keywords
        var keywords = new[] {
            // DML
            "SELECT", "FROM", "WHERE", "JOIN", "INNER", "LEFT", "RIGHT", "FULL", "OUTER",
            "ON", "AND", "OR", "NOT", "IN", "EXISTS", "BETWEEN", "LIKE", "IS", "NULL",
            "ORDER", "BY", "GROUP", "HAVING", "DISTINCT", "ALL", "AS", "UNION",
            // DDL
            "INSERT", "INTO", "UPDATE", "SET", "DELETE", "CREATE", "DROP", "ALTER", "TABLE",
            // Functions and aggregates
            "COUNT", "SUM", "AVG", "MIN", "MAX", "CASE", "WHEN", "THEN", "ELSE", "END",
            // Data types
            "INT", "INTEGER", "VARCHAR", "CHAR", "TEXT", "DATE", "DATETIME", "TIMESTAMP",
            "DECIMAL", "NUMERIC", "FLOAT", "DOUBLE", "BOOLEAN", "BOOL",
            // Other common keywords
            "WITH", "LIMIT", "OFFSET", "TOP", "FETCH", "FIRST", "ROWS", "ONLY", "DESC", "ASC"
        };

        foreach (var keyword in keywords)
        {
            formatted = System.Text.RegularExpressions.Regex.Replace(
                formatted,
                $@"\b{keyword}\b",
                $"<span class=\"sql-keyword\">{keyword}</span>",
                System.Text.RegularExpressions.RegexOptions.IgnoreCase
            );
        }

        // Highlight numbers (avoid matching inside HTML entities)
        formatted = System.Text.RegularExpressions.Regex.Replace(
            formatted,
            @"(?<!&[#\w])\b(\d+\.?\d*)\b(?![;\w])",
            "<span class=\"sql-number\">$1</span>"
        );

        // Highlight table/column names in brackets (already escaped as &lt; and &gt;)
        formatted = System.Text.RegularExpressions.Regex.Replace(
            formatted,
            @"\[([^\]]+)\]",
            "<span class=\"sql-identifier\">[$1]</span>"
        );

        formatted = System.Text.RegularExpressions.Regex.Replace(
            formatted,
            @"`([^`]+)`",
            "<span class=\"sql-identifier\">`$1`</span>"
        );

        // Restore string literals
        for (int i = 0; i < stringLiterals.Count; i++)
        {
            formatted = formatted.Replace(
                $"___STRING_LITERAL_{i}___",
                $"<span class=\"sql-string\">{EscapeHtml(stringLiterals[i])}</span>"
            );
        }

        // Restore comments
        for (int i = 0; i < comments.Count; i++)
        {
            formatted = formatted.Replace(
                $"___COMMENT_{i}___",
                $"<span class=\"sql-comment\">{EscapeHtml(comments[i])}</span>"
            );
        }

        return formatted;
    }

    private void AppendJavaScript()
    {
        _html.AppendLine("    <script>");
        _html.AppendLine(@"
        document.addEventListener('DOMContentLoaded', function() {
            // ═══════════════════════════════════════════════════════════════
            // INTERSECTION OBSERVER FOR SCROLL ANIMATIONS
            // ═══════════════════════════════════════════════════════════════

            const observerOptions = {
                root: null,
                rootMargin: '0px',
                threshold: 0.1
            };

            const fadeInObserver = new IntersectionObserver((entries) => {
                entries.forEach(entry => {
                    if (entry.isIntersecting) {
                        entry.target.style.opacity = '1';
                        entry.target.style.transform = 'translateY(0)';
                        fadeInObserver.unobserve(entry.target);
                    }
                });
            }, observerOptions);

            // Observe filter cards and dependency trees for scroll-triggered animations
            // Note: Tables excluded from fade-in to ensure visibility in all browsers
            document.querySelectorAll('.filter-info, .action-info, .dependency-tree').forEach(el => {
                el.style.opacity = '0';
                el.style.transform = 'translateY(20px)';
                el.style.transition = 'opacity 0.5s cubic-bezier(0.16, 1, 0.3, 1), transform 0.5s cubic-bezier(0.16, 1, 0.3, 1)';
                fadeInObserver.observe(el);
            });

            // ═══════════════════════════════════════════════════════════════
            // COLLAPSIBLE SECTIONS
            // ═══════════════════════════════════════════════════════════════

            const collapsibles = document.querySelectorAll('.collapsible-header');

            collapsibles.forEach(function(header) {
                header.addEventListener('click', function() {
                    const toggle = this.querySelector('.collapsible-toggle');
                    const content = this.nextElementSibling;

                    if (content && content.classList.contains('collapsible-content')) {
                        toggle.classList.toggle('collapsed');
                        content.classList.toggle('collapsed');
                    }
                });
            });

            // Collapse/Expand All buttons
            const collapseAllBtns = document.querySelectorAll('.collapse-all-btn');
            collapseAllBtns.forEach(function(btn) {
                btn.addEventListener('click', function() {
                    const section = this.closest('.section');
                    const isExpanding = this.textContent.includes('Expand');

                    const toggles = section.querySelectorAll('.collapsible-toggle');
                    const contents = section.querySelectorAll('.collapsible-content');

                    toggles.forEach(function(toggle, index) {
                        setTimeout(() => {
                            if (isExpanding) {
                                toggle.classList.remove('collapsed');
                            } else {
                                toggle.classList.add('collapsed');
                            }
                        }, index * 50);
                    });

                    contents.forEach(function(content, index) {
                        setTimeout(() => {
                            if (isExpanding) {
                                content.classList.remove('collapsed');
                            } else {
                                content.classList.add('collapsed');
                            }
                        }, index * 50);
                    });

                    this.textContent = isExpanding ? 'Collapse All' : 'Expand All';
                });
            });

            // ═══════════════════════════════════════════════════════════════
            // TOC COLLAPSIBLE SECTIONS
            // ═══════════════════════════════════════════════════════════════

            document.querySelectorAll('.toc-toggle').forEach(toggle => {
                toggle.addEventListener('click', function(e) {
                    e.stopPropagation();
                    const tocItem = this.closest('.toc-item');
                    if (tocItem) {
                        tocItem.classList.toggle('collapsed');
                    }
                });
            });

            // ═══════════════════════════════════════════════════════════════
            // SMOOTH SCROLL FOR TOC LINKS
            // ═══════════════════════════════════════════════════════════════

            document.querySelectorAll('.toc a[href^=""#""]').forEach(anchor => {
                anchor.addEventListener('click', function(e) {
                    e.preventDefault();
                    const target = document.querySelector(this.getAttribute('href'));
                    if (target) {
                        target.scrollIntoView({
                            behavior: 'smooth',
                            block: 'start'
                        });
                        // Update active state immediately on click
                        document.querySelectorAll('.toc a').forEach(a => a.classList.remove('active'));
                        this.classList.add('active');
                    }
                });
            });

            // ═══════════════════════════════════════════════════════════════
            // SCROLL SPY FOR TOC ACTIVE STATE
            // ═══════════════════════════════════════════════════════════════

            const tocLinks = document.querySelectorAll('.toc a[href^=""#""]');
            const sections = [];

            tocLinks.forEach(link => {
                const href = link.getAttribute('href');
                const section = document.querySelector(href);
                if (section) {
                    sections.push({ element: section, link: link });
                }
            });

            function updateActiveLink() {
                const scrollPosition = window.scrollY + 100; // Offset for better UX

                let activeSection = null;

                // Find the section that's currently in view
                for (let i = sections.length - 1; i >= 0; i--) {
                    const section = sections[i];
                    if (section.element.offsetTop <= scrollPosition) {
                        activeSection = section;
                        break;
                    }
                }

                // Update active class
                tocLinks.forEach(link => link.classList.remove('active'));
                if (activeSection) {
                    activeSection.link.classList.add('active');

                    // Scroll the TOC sidebar to keep active link visible
                    const sidebar = document.querySelector('.toc-sidebar');
                    const linkRect = activeSection.link.getBoundingClientRect();
                    const sidebarRect = sidebar.getBoundingClientRect();

                    if (linkRect.top < sidebarRect.top + 100 || linkRect.bottom > sidebarRect.bottom - 50) {
                        activeSection.link.scrollIntoView({ behavior: 'smooth', block: 'center' });
                    }
                }
            }

            // Throttle scroll event for performance
            let scrollTimeout;
            window.addEventListener('scroll', function() {
                if (scrollTimeout) {
                    window.cancelAnimationFrame(scrollTimeout);
                }
                scrollTimeout = window.requestAnimationFrame(updateActiveLink);
            });

            // Initial call to set active state
            updateActiveLink();

            // ═══════════════════════════════════════════════════════════════
            // STAT CARD NUMBER ANIMATION
            // ═══════════════════════════════════════════════════════════════

            const statObserver = new IntersectionObserver((entries) => {
                entries.forEach(entry => {
                    if (entry.isIntersecting) {
                        const valueEl = entry.target;
                        const finalValue = parseInt(valueEl.textContent);
                        if (!isNaN(finalValue) && finalValue > 0) {
                            animateNumber(valueEl, 0, finalValue, 800);
                        }
                        statObserver.unobserve(entry.target);
                    }
                });
            }, { threshold: 0.5 });

            document.querySelectorAll('.stat-value').forEach(el => {
                statObserver.observe(el);
            });

            function animateNumber(el, start, end, duration) {
                const startTime = performance.now();
                const easeOutQuart = t => 1 - Math.pow(1 - t, 4);

                function update(currentTime) {
                    const elapsed = currentTime - startTime;
                    const progress = Math.min(elapsed / duration, 1);
                    const easedProgress = easeOutQuart(progress);
                    const current = Math.round(start + (end - start) * easedProgress);
                    el.textContent = current;

                    if (progress < 1) {
                        requestAnimationFrame(update);
                    }
                }

                requestAnimationFrame(update);
            }

            // ═══════════════════════════════════════════════════════════════
            // TABLE ROW HOVER EFFECTS
            // ═══════════════════════════════════════════════════════════════

            document.querySelectorAll('tbody tr').forEach(row => {
                row.addEventListener('mouseenter', function() {
                    this.style.transition = 'background-color 0.2s ease';
                });
            });
        });
        ");
        _html.AppendLine("    </script>");
    }

    /// <summary>
    /// Append field lineage information showing the path from Tableau field to data source
    /// </summary>
    private void AppendFieldLineage(Core.Models.FieldLineage? lineage, string indent = "            ", List<string>? allowedValues = null)
    {
        if (lineage == null)
            return;

        _html.AppendLine($"{indent}<div class=\"field-lineage\">");
        _html.AppendLine($"{indent}    <p>Field lineage</p>");
        _html.AppendLine($"{indent}    <div class=\"table-wrapper\">");
        _html.AppendLine($"{indent}    <table class=\"lineage-table\">");

        // Display Name (Tableau Name)
        _html.AppendLine($"{indent}        <tr>");
        _html.AppendLine($"{indent}            <td class=\"lineage-label\">Tableau Field:</td>");
        _html.AppendLine($"{indent}            <td class=\"lineage-value\">{EscapeHtml(lineage.DisplayName)}</td>");
        _html.AppendLine($"{indent}        </tr>");

        // Derivation (if any)
        if (!string.IsNullOrEmpty(lineage.Derivation))
        {
            _html.AppendLine($"{indent}        <tr>");
            _html.AppendLine($"{indent}            <td class=\"lineage-label\">Derivation:</td>");
            _html.AppendLine($"{indent}            <td class=\"lineage-value\">{EscapeHtml(lineage.Derivation)}</td>");
            _html.AppendLine($"{indent}        </tr>");
        }

        // Base Field Name - with link to data source field if available
        if (!string.IsNullOrEmpty(lineage.BaseFieldName) && lineage.BaseFieldName != lineage.DisplayName)
        {
            _html.AppendLine($"{indent}        <tr>");
            _html.AppendLine($"{indent}            <td class=\"lineage-label\">Base Field:</td>");

            // Try to create a link to the field in the data sources section
            var baseFieldLink = GetFieldLink(lineage.BaseFieldName, lineage.DataSourceName);
            if (!string.IsNullOrEmpty(baseFieldLink))
            {
                _html.AppendLine($"{indent}            <td class=\"lineage-value\"><a href=\"{baseFieldLink}\">{EscapeHtml(lineage.BaseFieldName)}</a></td>");
            }
            else
            {
                _html.AppendLine($"{indent}            <td class=\"lineage-value\">{EscapeHtml(lineage.BaseFieldName)}</td>");
            }
            _html.AppendLine($"{indent}        </tr>");
        }

        // Data Source - with link to parameters section if it's a parameter
        if (!string.IsNullOrEmpty(lineage.DataSourceName))
        {
            _html.AppendLine($"{indent}        <tr>");
            _html.AppendLine($"{indent}            <td class=\"lineage-label\">Data Source:</td>");

            if (lineage.DataSourceName == "Parameters" && !string.IsNullOrEmpty(lineage.BaseFieldName))
            {
                // Link to the parameter in the parameters section using the actual parameter name
                var paramId = $"param-{ToHtmlId(lineage.BaseFieldName)}";
                _html.AppendLine($"{indent}            <td class=\"lineage-value\"><a href=\"#{paramId}\">{EscapeHtml(lineage.DataSourceName)}</a></td>");
            }
            else
            {
                _html.AppendLine($"{indent}            <td class=\"lineage-value\">{EscapeHtml(lineage.DataSourceName)}</td>");
            }
            _html.AppendLine($"{indent}        </tr>");
        }

        // Allowed values (only for parameter-type filters with a distinct list)
        if (lineage.DataSourceName == "Parameters" && allowedValues != null && allowedValues.Any())
        {
            _html.AppendLine($"{indent}        <tr>");
            _html.AppendLine($"{indent}            <td class=\"lineage-label\">Allowed Values:</td>");
            _html.AppendLine($"{indent}            <td class=\"lineage-value\">{EscapeHtml(string.Join(", ", allowedValues))}</td>");
            _html.AppendLine($"{indent}        </tr>");
        }

        // Data Type and Role
        var typeInfo = new List<string>();
        if (!string.IsNullOrEmpty(lineage.DataType))
            typeInfo.Add(lineage.DataType);
        if (!string.IsNullOrEmpty(lineage.Role))
            typeInfo.Add(lineage.Role);

        if (typeInfo.Any())
        {
            _html.AppendLine($"{indent}        <tr>");
            _html.AppendLine($"{indent}            <td class=\"lineage-label\">Type:</td>");
            _html.AppendLine($"{indent}            <td class=\"lineage-value\">{EscapeHtml(string.Join(" / ", typeInfo))}</td>");
            _html.AppendLine($"{indent}        </tr>");
        }

        // Physical mapping (for non-calculated fields)
        if (!lineage.IsCalculated && !string.IsNullOrEmpty(lineage.PhysicalColumn))
        {
            var physicalRef = !string.IsNullOrEmpty(lineage.PhysicalTable)
                ? $"{lineage.PhysicalTable}.{lineage.PhysicalColumn}"
                : lineage.PhysicalColumn;

            _html.AppendLine($"{indent}        <tr>");
            _html.AppendLine($"{indent}            <td class=\"lineage-label\">Data Element:</td>");
            _html.AppendLine($"{indent}            <td class=\"lineage-value\"><code>{EscapeHtml(physicalRef)}</code></td>");
            _html.AppendLine($"{indent}        </tr>");
        }

        // For calculated fields, show formula and dependencies
        if (lineage.IsCalculated)
        {
            _html.AppendLine($"{indent}        <tr>");
            _html.AppendLine($"{indent}            <td class=\"lineage-label\">Calculation:</td>");
            _html.AppendLine($"{indent}            <td class=\"lineage-value\"><span class=\"badge badge-calculated\">Calculated Field</span></td>");
            _html.AppendLine($"{indent}        </tr>");

            if (!string.IsNullOrEmpty(lineage.Formula))
            {
                _html.AppendLine($"{indent}        <tr>");
                _html.AppendLine($"{indent}            <td class=\"lineage-label\">Formula:</td>");
                _html.AppendLine($"{indent}            <td class=\"lineage-value\"><code>{EscapeHtml(lineage.Formula)}</code></td>");
                _html.AppendLine($"{indent}        </tr>");
            }

            if (lineage.Dependencies != null && lineage.Dependencies.Any())
            {
                _html.AppendLine($"{indent}        <tr>");
                _html.AppendLine($"{indent}            <td class=\"lineage-label\">Depends On:</td>");
                _html.AppendLine($"{indent}            <td class=\"lineage-value\">{EscapeHtml(string.Join(", ", lineage.Dependencies))}</td>");
                _html.AppendLine($"{indent}        </tr>");
            }
        }

        _html.AppendLine($"{indent}    </table>");
        _html.AppendLine($"{indent}    </div>");
        _html.AppendLine($"{indent}</div>");
    }

    /// <summary>
    /// Append dashboard color theme information with color swatches
    /// </summary>
    private void AppendDashboardColorTheme(Core.Models.DashboardColorTheme colorTheme)
    {
        _html.Append("        <p class=\"color-theme-info\"><strong>Colors:</strong> ");

        var colorItems = new List<string>();

        // Primary color
        if (!string.IsNullOrEmpty(colorTheme.PrimaryColor))
        {
            colorItems.Add($"<span class=\"color-swatch\" style=\"background-color:{colorTheme.PrimaryColor}\" title=\"Primary: {colorTheme.PrimaryColor}\"></span> {colorTheme.PrimaryColor}");
        }

        // Mark colors
        if (colorTheme.MarkColors?.Any() == true)
        {
            foreach (var markColor in colorTheme.MarkColors)
            {
                colorItems.Add($"<span class=\"color-swatch\" style=\"background-color:{markColor}\" title=\"Mark: {markColor}\"></span> {markColor}");
            }
        }

        _html.Append(string.Join(" ", colorItems));

        // Palette name (append after colors if present)
        if (!string.IsNullOrEmpty(colorTheme.PaletteName))
        {
            _html.Append($" (Palette: {EscapeHtml(colorTheme.PaletteName)})");
        }

        _html.AppendLine("</p>");
    }

    /// <summary>
    /// Append KPI fields as a table with FIELD, DESCRIPTION, and COLOR columns
    /// </summary>
    private void AppendKpiFieldsTable(Core.Models.CustomizedLabel customizedLabel)
    {
        _html.AppendLine("                <div class=\"table-wrapper\">");
        _html.AppendLine("                <table class=\"kpi-fields-table\">");
        _html.AppendLine("                <thead>");
        _html.AppendLine("                    <tr>");
        _html.AppendLine("                        <th>Field</th>");
        _html.AppendLine("                        <th>Description</th>");
        _html.AppendLine("                        <th>Color</th>");
        _html.AppendLine("                    </tr>");
        _html.AppendLine("                </thead>");
        _html.AppendLine("                <tbody>");

        foreach (var fieldRole in customizedLabel.FieldRoles)
        {
            _html.AppendLine("                    <tr>");
            if (fieldRole.IsStaticText)
            {
                // Static text is displayed in italics
                _html.AppendLine($"                        <td><em>{EscapeHtml(fieldRole.FieldName)}</em></td>");
            }
            else
            {
                _html.AppendLine($"                        <td>{EscapeHtml(fieldRole.FieldName)}</td>");
            }
            _html.AppendLine($"                        <td>{EscapeHtml(fieldRole.Role)}</td>");

            // Color column with swatch
            if (!string.IsNullOrEmpty(fieldRole.FontColor))
            {
                var inheritedIndicator = fieldRole.IsInherited ? " <em>(inherited)</em>" : "";
                var title = fieldRole.IsInherited ? $"{fieldRole.FontColor} (inherited from worksheet default)" : fieldRole.FontColor;
                _html.AppendLine($"                        <td><span class=\"color-swatch\" style=\"background-color:{fieldRole.FontColor}\" title=\"{title}\"></span> {fieldRole.FontColor}{inheritedIndicator}</td>");
            }
            else
            {
                _html.AppendLine("                        <td></td>");
            }
            _html.AppendLine("                    </tr>");
        }

        _html.AppendLine("                </tbody>");
        _html.AppendLine("                </table>");
        _html.AppendLine("                </div>");
    }

    /// <summary>
    /// Append customized label information for worksheets (text displays with mixed field references and static text)
    /// </summary>
    private void AppendCustomizedLabelForWorksheet(Core.Models.CustomizedLabel customizedLabel)
    {
        _html.AppendLine("            <h4>Display Label</h4>");
        _html.AppendLine("            <div class=\"table-wrapper\">");
        _html.AppendLine("            <table class=\"kpi-fields-table\">");
        _html.AppendLine("            <thead>");
        _html.AppendLine("                <tr>");
        _html.AppendLine("                    <th>Element</th>");
        _html.AppendLine("                    <th>Type</th>");
        _html.AppendLine("                    <th>Color</th>");
        _html.AppendLine("                </tr>");
        _html.AppendLine("            </thead>");
        _html.AppendLine("            <tbody>");

        foreach (var fieldRole in customizedLabel.FieldRoles)
        {
            _html.AppendLine("                <tr>");
            if (fieldRole.IsStaticText)
            {
                // Static text is displayed in italics
                _html.AppendLine($"                    <td><em>{EscapeHtml(fieldRole.FieldName)}</em></td>");
                _html.AppendLine("                    <td>Static Text</td>");
            }
            else
            {
                _html.AppendLine($"                    <td>{EscapeHtml(fieldRole.FieldName)}</td>");
                _html.AppendLine($"                    <td>{EscapeHtml(fieldRole.Role)}</td>");
            }

            // Color column with swatch
            if (!string.IsNullOrEmpty(fieldRole.FontColor))
            {
                var inheritedIndicator = fieldRole.IsInherited ? " <em>(inherited)</em>" : "";
                var title = fieldRole.IsInherited ? $"{fieldRole.FontColor} (inherited from worksheet default)" : fieldRole.FontColor;
                _html.AppendLine($"                    <td><span class=\"color-swatch\" style=\"background-color:{fieldRole.FontColor}\" title=\"{title}\"></span> {fieldRole.FontColor}{inheritedIndicator}</td>");
            }
            else
            {
                _html.AppendLine("                    <td></td>");
            }
            _html.AppendLine("                </tr>");
        }

        _html.AppendLine("            </tbody>");
        _html.AppendLine("            </table>");
        _html.AppendLine("            </div>");
    }

    /// <summary>
    /// Append worksheet-level filters showing what filter conditions are applied
    /// </summary>
    private void AppendWorksheetFilters(List<Core.Models.Filter>? filters)
    {
        if (filters == null || !filters.Any())
            return;

        // Filter out action filters and internal filters
        var displayableFilters = filters
            .Where(f => !string.IsNullOrEmpty(f.Field) &&
                       !f.Field.StartsWith("Action (") &&
                       !f.Field.StartsWith("[Action ("))
            .ToList();

        if (!displayableFilters.Any())
            return;

        _html.AppendLine("            <h4>Filters Applied</h4>");
        _html.AppendLine("            <div class=\"table-wrapper\">");
        _html.AppendLine("            <table>");
        _html.AppendLine("            <thead>");
        _html.AppendLine("                <tr>");
        _html.AppendLine("                    <th>Field</th>");
        _html.AppendLine("                    <th>Type</th>");
        _html.AppendLine("                    <th>Value(s)</th>");
        _html.AppendLine("                </tr>");
        _html.AppendLine("            </thead>");
        _html.AppendLine("            <tbody>");

        foreach (var filter in displayableFilters)
        {
            _html.AppendLine("                <tr>");
            _html.AppendLine($"                    <td>{EscapeHtml(filter.Field)}</td>");
            _html.AppendLine($"                    <td>{EscapeHtml(filter.FilterType)}</td>");

            // Show the filter value(s)
            var valueDisplay = GetFilterValueDisplay(filter);
            _html.AppendLine($"                    <td>{EscapeHtml(valueDisplay)}</td>");
            _html.AppendLine("                </tr>");
        }

        _html.AppendLine("            </tbody>");
        _html.AppendLine("            </table>");
        _html.AppendLine("            </div>");
    }

    /// <summary>
    /// Get a displayable string for the filter value(s)
    /// </summary>
    private string GetFilterValueDisplay(Core.Models.Filter filter)
    {
        // For quantitative filters with range
        if (filter.Type == "quantitative")
        {
            if (filter.Max != null && filter.Min == null)
            {
                return $"≤ {filter.Max}";
            }
            else if (filter.Min != null && filter.Max == null)
            {
                return $"≥ {filter.Min}";
            }
            else if (filter.Min != null && filter.Max != null)
            {
                return $"{filter.Min} to {filter.Max}";
            }
        }

        // For categorical filters with default selection
        if (!string.IsNullOrEmpty(filter.DefaultSelection))
        {
            return filter.DefaultSelection;
        }

        // For date filters
        if (filter.Type.Contains("date") && !string.IsNullOrEmpty(filter.Period))
        {
            var periodDesc = filter.Period;
            if (filter.PeriodOffset.HasValue)
            {
                periodDesc += filter.PeriodOffset > 0 ? $" +{filter.PeriodOffset}" : $" {filter.PeriodOffset}";
            }
            return periodDesc;
        }

        return "(Not specified)";
    }

    /// <summary>
    /// Append mark encoding information (Color, Size, Text, Shape, etc.)
    /// </summary>
    /// <param name="encodings">Mark encodings to display</param>
    /// <param name="hasActualTooltipContent">If true, label as "Tooltip"; if false, label as "Detail" since it's just shelf placement</param>
    private void AppendMarkEncodings(Core.Models.MarkEncodings? encodings, bool hasActualTooltipContent = false)
    {
        if (encodings == null)
            return;

        var hasAnyEncoding = encodings.Color != null ||
                             encodings.Size != null ||
                             (encodings.Text?.Fields.Any() == true) ||
                             encodings.Shape != null ||
                             encodings.DetailFields.Any() ||
                             encodings.TooltipFields.Any();

        if (!hasAnyEncoding)
            return;

        _html.AppendLine("                <div class=\"mark-encodings\">");
        _html.AppendLine("                    <h6>Marks</h6>");
        _html.AppendLine("                    <div class=\"table-wrapper\">");
        _html.AppendLine("                    <table class=\"marks-table\">");
        _html.AppendLine("                    <thead>");
        _html.AppendLine("                        <tr>");
        _html.AppendLine("                            <th>Shelf</th>");
        _html.AppendLine("                            <th>Field</th>");
        _html.AppendLine("                            <th>Details</th>");
        _html.AppendLine("                        </tr>");
        _html.AppendLine("                    </thead>");
        _html.AppendLine("                    <tbody>");

        // Color encoding
        if (encodings.Color != null)
        {
            _html.AppendLine("                        <tr>");
            _html.AppendLine("                            <td><span class=\"mark-shelf color-shelf\">Color</span></td>");

            // Field column - show field name or "—" if no field
            if (!string.IsNullOrEmpty(encodings.Color.Field))
            {
                _html.AppendLine($"                            <td>{EscapeHtml(encodings.Color.Field)}</td>");
            }
            else
            {
                _html.AppendLine("                            <td>—</td>");
            }

            // Details column - always show color here (mark color, palette info, etc.)
            var colorDetailsList = new List<string>();

            // Mark color with swatch
            // Skip mark-color if categorical color mappings exist (they take precedence)
            if (!string.IsNullOrEmpty(encodings.Color.MarkColor) && !encodings.Color.ColorMappings.Any())
            {
                var markColorSwatch = $"<span class=\"color-swatch\" style=\"background-color:{encodings.Color.MarkColor}\" title=\"{(string.IsNullOrEmpty(encodings.Color.Field) ? "Static color - all marks use this single color" : encodings.Color.MarkColor)}\"></span>";
                if (string.IsNullOrEmpty(encodings.Color.Field))
                {
                    // No field - this is a static color
                    colorDetailsList.Add($"{markColorSwatch} {encodings.Color.MarkColor} <em title=\"Static color - all marks use this single color\">(static)</em>");
                }
                else
                {
                    // Field with mark color
                    colorDetailsList.Add($"{markColorSwatch} {encodings.Color.MarkColor}");
                }
            }

            // Color mappings — group by CASE parameter branches if applicable
            if (encodings.Color.ColorMappings.Any())
            {
                var grouped = BuildGroupedColorMappings(encodings.Color);
                if (grouped != null)
                {
                    foreach (var group in grouped)
                    {
                        colorDetailsList.Add($"<em style=\"font-size:0.82em;color:#6b7280;display:block;margin-top:4px\">{EscapeHtml(group.Key)}</em>");
                        foreach (var mapping in group.Value)
                        {
                            var colorSwatch = $"<span class=\"color-swatch\" style=\"background-color:{mapping.Value}\" title=\"{EscapeHtml(mapping.Key)}: {mapping.Value}\"></span>";
                            var displayKey = mapping.Key == "%null%" ? "%null% <em>(null/missing values)</em>" : EscapeHtml(mapping.Key);
                            colorDetailsList.Add($"{colorSwatch} <code>{mapping.Value}</code> - {displayKey}");
                        }
                    }
                }
                else
                {
                    foreach (var mapping in encodings.Color.ColorMappings)
                    {
                        var colorSwatch = $"<span class=\"color-swatch\" style=\"background-color:{mapping.Value}\" title=\"{EscapeHtml(mapping.Key)}: {mapping.Value}\"></span>";
                        var displayKey = mapping.Key == "%null%" ? "%null% <em>(null/missing values)</em>" : EscapeHtml(mapping.Key);
                        colorDetailsList.Add($"{colorSwatch} <code>{mapping.Value}</code> - {displayKey}");
                    }
                }
            }

            // Palette name
            if (!string.IsNullOrEmpty(encodings.Color.PaletteName))
            {
                colorDetailsList.Add($"Palette: {EscapeHtml(encodings.Color.PaletteName)}");
            }

            // Output details
            if (colorDetailsList.Any())
            {
                _html.AppendLine($"                            <td>{string.Join("<br>", colorDetailsList)}</td>");
            }
            else if (!string.IsNullOrEmpty(encodings.Color.PaletteType) && encodings.Color.PaletteType != "column")
            {
                _html.AppendLine($"                            <td>{EscapeHtml(encodings.Color.PaletteType)}</td>");
            }
            else if (!string.IsNullOrEmpty(encodings.Color.Field))
            {
                // Field on Color shelf but no explicit colors defined - uses Tableau's default palette
                _html.AppendLine("                            <td>Default palette</td>");
            }
            else
            {
                _html.AppendLine("                            <td></td>");
            }
            _html.AppendLine("                        </tr>");
        }

        // Size encoding
        if (encodings.Size != null)
        {
            _html.AppendLine("                        <tr>");
            _html.AppendLine("                            <td><span class=\"mark-shelf size-shelf\">Size</span></td>");
            _html.AppendLine($"                            <td>{EscapeHtml(encodings.Size.Field)}</td>");
            _html.AppendLine($"                            <td>{(!string.IsNullOrEmpty(encodings.Size.Aggregation) && encodings.Size.Aggregation != "None" ? EscapeHtml(encodings.Size.Aggregation) : "")}</td>");
            _html.AppendLine("                        </tr>");
        }

        // Text/Label encoding
        if (encodings.Text?.Fields.Any() == true)
        {
            _html.AppendLine("                        <tr>");
            _html.AppendLine("                            <td><span class=\"mark-shelf text-shelf\">Text</span></td>");
            _html.AppendLine($"                            <td>{EscapeHtml(string.Join(", ", encodings.Text.Fields))}</td>");
            _html.AppendLine("                            <td>Labels</td>");
            _html.AppendLine("                        </tr>");
        }

        // Shape encoding
        // Skip default shape values (starting with colon like ':filled/circle') when no field is assigned
        if (encodings.Shape != null &&
            !string.IsNullOrEmpty(encodings.Shape.ShapeType) &&
            (!encodings.Shape.ShapeType.StartsWith(":") || !string.IsNullOrEmpty(encodings.Shape.Field)))
        {
            _html.AppendLine("                        <tr>");
            _html.AppendLine("                            <td><span class=\"mark-shelf shape-shelf\">Shape</span></td>");
            _html.AppendLine($"                            <td>{(string.IsNullOrEmpty(encodings.Shape.Field) ? "-" : EscapeHtml(encodings.Shape.Field))}</td>");
            _html.AppendLine($"                            <td>{EscapeHtml(encodings.Shape.ShapeType)}</td>");
            _html.AppendLine("                        </tr>");
        }

        // Detail encoding
        if (encodings.DetailFields.Any())
        {
            _html.AppendLine("                        <tr>");
            _html.AppendLine("                            <td><span class=\"mark-shelf detail-shelf\">Detail</span></td>");
            _html.AppendLine($"                            <td>{EscapeHtml(string.Join(", ", encodings.DetailFields))}</td>");
            _html.AppendLine("                            <td>Adds detail to marks</td>");
            _html.AppendLine("                        </tr>");
        }

        // Tooltip/Detail encoding - only call it "Tooltip" if there's actual hover tooltip content
        if (encodings.TooltipFields.Any())
        {
            // If no actual tooltip content exists, these are just "Detail" fields placed on the tooltip shelf
            var shelfName = hasActualTooltipContent ? "Tooltip" : "Detail";
            var shelfClass = hasActualTooltipContent ? "tooltip-shelf" : "detail-shelf";
            var description = hasActualTooltipContent ? "Shows on hover" : "Mark encoding (no tooltip display)";

            _html.AppendLine("                        <tr>");
            _html.AppendLine($"                            <td><span class=\"mark-shelf {shelfClass}\">{shelfName}</span></td>");
            _html.AppendLine($"                            <td>{EscapeHtml(string.Join(", ", encodings.TooltipFields))}</td>");
            _html.AppendLine($"                            <td>{description}</td>");
            _html.AppendLine("                        </tr>");
        }

        _html.AppendLine("                    </tbody>");
        _html.AppendLine("                    </table>");
        _html.AppendLine("                    </div>");
        _html.AppendLine("                </div>");
    }

    /// <summary>
    /// Truncate a value for display purposes
    /// </summary>
    /// <summary>
    /// If the color field is a CASE-based calculated field driven by a parameter,
    /// groups color mappings by the parameter's alias labels (e.g., "Race", "Sex", "Age").
    /// Returns null if no grouping applies.
    /// </summary>
    /// <summary>
    /// Returns pre-built GroupedColorMappings from the parser (already keyed with friendly names).
    /// Returns null if no grouping applies.
    /// </summary>
    private Dictionary<string, Dictionary<string, string>>? BuildGroupedColorMappings(
        Core.Models.ColorEncoding color)
    {
        if (color.GroupedColorMappings == null || color.GroupedColorMappings.Count <= 1)
            return null;

        return color.GroupedColorMappings;
    }

    private string TruncateValue(string value, int maxLength)
    {
        if (string.IsNullOrEmpty(value) || value.Length <= maxLength)
            return value;
        return value.Substring(0, maxLength - 3) + "...";
    }

    /// <summary>
    /// Append mark encoding information for worksheets (different indentation)
    /// </summary>
    private void AppendMarkEncodingsForWorksheet(Core.Models.MarkEncodings? encodings)
    {
        if (encodings == null)
            return;

        var hasAnyEncoding = encodings.Color != null ||
                             encodings.Size != null ||
                             (encodings.Text?.Fields.Any() == true) ||
                             encodings.Shape != null ||
                             encodings.DetailFields.Any() ||
                             encodings.TooltipFields.Any();

        if (!hasAnyEncoding)
            return;

        _html.AppendLine("            <div class=\"mark-encodings\">");
        _html.AppendLine("                <h4>Marks</h4>");
        _html.AppendLine("                <div class=\"table-wrapper\">");
        _html.AppendLine("                <table class=\"marks-table\">");
        _html.AppendLine("                <thead>");
        _html.AppendLine("                    <tr>");
        _html.AppendLine("                        <th>Shelf</th>");
        _html.AppendLine("                        <th>Field</th>");
        _html.AppendLine("                        <th>Details</th>");
        _html.AppendLine("                    </tr>");
        _html.AppendLine("                </thead>");
        _html.AppendLine("                <tbody>");

        // Color encoding
        if (encodings.Color != null)
        {
            _html.AppendLine("                    <tr>");
            _html.AppendLine("                        <td><span class=\"mark-shelf color-shelf\">Color</span></td>");

            // Field column - show field name or "—" if no field
            if (!string.IsNullOrEmpty(encodings.Color.Field))
            {
                _html.AppendLine($"                        <td>{EscapeHtml(encodings.Color.Field)}</td>");
            }
            else
            {
                _html.AppendLine("                        <td>—</td>");
            }

            // Details column - always show color here (mark color, palette info, etc.)
            var colorDetailsList = new List<string>();

            // Mark color with swatch
            // Skip mark-color if categorical color mappings exist (they take precedence)
            if (!string.IsNullOrEmpty(encodings.Color.MarkColor) && !encodings.Color.ColorMappings.Any())
            {
                var markColorSwatch = $"<span class=\"color-swatch\" style=\"background-color:{encodings.Color.MarkColor}\" title=\"{(string.IsNullOrEmpty(encodings.Color.Field) ? "Static color - all marks use this single color" : encodings.Color.MarkColor)}\"></span>";
                if (string.IsNullOrEmpty(encodings.Color.Field))
                {
                    // No field - this is a static color
                    colorDetailsList.Add($"{markColorSwatch} {encodings.Color.MarkColor} <em title=\"Static color - all marks use this single color\">(static)</em>");
                }
                else
                {
                    // Field with mark color
                    colorDetailsList.Add($"{markColorSwatch} {encodings.Color.MarkColor}");
                }
            }

            // Color mappings — group by CASE parameter branches if applicable
            if (encodings.Color.ColorMappings.Any())
            {
                var grouped = BuildGroupedColorMappings(encodings.Color);
                if (grouped != null)
                {
                    foreach (var group in grouped)
                    {
                        colorDetailsList.Add($"<em style=\"font-size:0.82em;color:#6b7280;display:block;margin-top:4px\">{EscapeHtml(group.Key)}</em>");
                        foreach (var mapping in group.Value)
                        {
                            var colorSwatch = $"<span class=\"color-swatch\" style=\"background-color:{mapping.Value}\" title=\"{EscapeHtml(mapping.Key)}: {mapping.Value}\"></span>";
                            var displayKey = mapping.Key == "%null%" ? "%null% <em>(null/missing values)</em>" : EscapeHtml(mapping.Key);
                            colorDetailsList.Add($"{colorSwatch} {displayKey}");
                        }
                    }
                }
                else
                {
                    foreach (var mapping in encodings.Color.ColorMappings)
                    {
                        var colorSwatch = $"<span class=\"color-swatch\" style=\"background-color:{mapping.Value}\" title=\"{EscapeHtml(mapping.Key)}: {mapping.Value}\"></span>";
                        var displayKey = mapping.Key == "%null%" ? "%null% <em>(null/missing values)</em>" : EscapeHtml(mapping.Key);
                        colorDetailsList.Add($"{colorSwatch} {displayKey}");
                    }
                }
            }

            // Palette name
            if (!string.IsNullOrEmpty(encodings.Color.PaletteName))
            {
                colorDetailsList.Add($"Palette: {EscapeHtml(encodings.Color.PaletteName)}");
            }

            // Output details
            if (colorDetailsList.Any())
            {
                _html.AppendLine($"                        <td>{string.Join("<br>", colorDetailsList)}</td>");
            }
            else if (!string.IsNullOrEmpty(encodings.Color.PaletteType) && encodings.Color.PaletteType != "column")
            {
                _html.AppendLine($"                        <td>{EscapeHtml(encodings.Color.PaletteType)}</td>");
            }
            else if (!string.IsNullOrEmpty(encodings.Color.Field))
            {
                // Field on Color shelf but no explicit colors defined - uses Tableau's default palette
                _html.AppendLine("                        <td>Default palette</td>");
            }
            else
            {
                _html.AppendLine("                        <td></td>");
            }
            _html.AppendLine("                    </tr>");
        }

        // Size encoding
        if (encodings.Size != null)
        {
            _html.AppendLine("                    <tr>");
            _html.AppendLine("                        <td><span class=\"mark-shelf size-shelf\">Size</span></td>");
            _html.AppendLine($"                        <td>{EscapeHtml(encodings.Size.Field)}</td>");
            _html.AppendLine($"                        <td>{(!string.IsNullOrEmpty(encodings.Size.Aggregation) && encodings.Size.Aggregation != "None" ? EscapeHtml(encodings.Size.Aggregation) : "")}</td>");
            _html.AppendLine("                    </tr>");
        }

        // Text/Label encoding
        if (encodings.Text?.Fields.Any() == true)
        {
            _html.AppendLine("                    <tr>");
            _html.AppendLine("                        <td><span class=\"mark-shelf text-shelf\">Text</span></td>");
            _html.AppendLine($"                        <td>{EscapeHtml(string.Join(", ", encodings.Text.Fields))}</td>");
            _html.AppendLine("                        <td>Labels</td>");
            _html.AppendLine("                    </tr>");
        }

        // Shape encoding
        if (encodings.Shape != null)
        {
            _html.AppendLine("                    <tr>");
            _html.AppendLine("                        <td><span class=\"mark-shelf shape-shelf\">Shape</span></td>");
            _html.AppendLine($"                        <td>{(string.IsNullOrEmpty(encodings.Shape.Field) ? "-" : EscapeHtml(encodings.Shape.Field))}</td>");
            _html.AppendLine($"                        <td>{EscapeHtml(encodings.Shape.ShapeType)}</td>");
            _html.AppendLine("                    </tr>");
        }

        // Detail encoding
        if (encodings.DetailFields.Any())
        {
            _html.AppendLine("                    <tr>");
            _html.AppendLine("                        <td><span class=\"mark-shelf detail-shelf\">Detail</span></td>");
            _html.AppendLine($"                        <td>{EscapeHtml(string.Join(", ", encodings.DetailFields))}</td>");
            _html.AppendLine("                        <td>Adds detail to marks</td>");
            _html.AppendLine("                    </tr>");
        }

        // Tooltip encoding
        if (encodings.TooltipFields.Any())
        {
            _html.AppendLine("                    <tr>");
            _html.AppendLine("                        <td><span class=\"mark-shelf tooltip-shelf\">Tooltip</span></td>");
            _html.AppendLine($"                        <td>{EscapeHtml(string.Join(", ", encodings.TooltipFields))}</td>");
            _html.AppendLine("                        <td></td>");
            _html.AppendLine("                    </tr>");
        }

        _html.AppendLine("                </tbody>");
        _html.AppendLine("                </table>");
        _html.AppendLine("                </div>");
        _html.AppendLine("            </div>");
    }

    /// <summary>
    /// Append map configuration information for worksheets
    /// </summary>
    private void AppendMapConfigurationForWorksheet(Core.Models.MapConfiguration? mapConfig)
    {
        if (mapConfig == null)
            return;

        _html.AppendLine("            <div class=\"map-configuration\">");
        _html.AppendLine("                <h4>Map Configuration</h4>");

        // Map Source
        if (!string.IsNullOrEmpty(mapConfig.MapSource))
        {
            _html.AppendLine($"                <p><strong>Map Provider:</strong> {EscapeHtml(mapConfig.MapSource)}</p>");
        }

        // Geographic Fields
        if (mapConfig.GeographicFields.Any())
        {
            _html.AppendLine("                <h5>Geographic Fields</h5>");
            _html.AppendLine("                <div class=\"table-wrapper\">");
            _html.AppendLine("                <table class=\"geographic-fields-table\">");
            _html.AppendLine("                <thead>");
            _html.AppendLine("                    <tr>");
            _html.AppendLine("                        <th>Field</th>");
            _html.AppendLine("                        <th>Role</th>");
            _html.AppendLine("                        <th>Shelf</th>");
            _html.AppendLine("                    </tr>");
            _html.AppendLine("                </thead>");
            _html.AppendLine("                <tbody>");

            foreach (var geoField in mapConfig.GeographicFields)
            {
                var fieldDisplay = geoField.FieldName;
                if (geoField.IsGenerated)
                {
                    fieldDisplay += " <em>(generated)</em>";
                }

                _html.AppendLine("                    <tr>");
                _html.AppendLine($"                        <td>{EscapeHtml(geoField.FieldName)}{(geoField.IsGenerated ? " <em>(generated)</em>" : "")}</td>");
                _html.AppendLine($"                        <td>{EscapeHtml(geoField.GeographicRole)}</td>");
                _html.AppendLine($"                        <td>{EscapeHtml(geoField.Shelf)}</td>");
                _html.AppendLine("                    </tr>");
            }

            _html.AppendLine("                </tbody>");
            _html.AppendLine("                </table>");
            _html.AppendLine("                </div>");
        }

        // Geometry Encoding
        if (mapConfig.HasGeometryEncoding)
        {
            _html.AppendLine("                <p><strong>Geometry Encoding:</strong> ");
            if (!string.IsNullOrEmpty(mapConfig.GeometryColumn))
            {
                _html.AppendLine($"{EscapeHtml(mapConfig.GeometryColumn)}</p>");
            }
            else
            {
                _html.AppendLine("Yes</p>");
            }
        }

        // Map Styling
        var styleInfo = new List<string>();
        if (!string.IsNullOrEmpty(mapConfig.BaseMapStyle))
        {
            styleInfo.Add($"Style: {EscapeHtml(mapConfig.BaseMapStyle)}");
        }
        if (mapConfig.Washout.HasValue)
        {
            var washoutPercent = (int)(mapConfig.Washout.Value * 100);
            styleInfo.Add($"Washout: {washoutPercent}%");
        }
        if (mapConfig.ShowLabels.HasValue)
        {
            styleInfo.Add($"Labels: {(mapConfig.ShowLabels.Value ? "Shown" : "Hidden")}");
        }

        if (styleInfo.Any())
        {
            _html.AppendLine($"                <p><strong>Map Style:</strong> {string.Join(", ", styleInfo)}</p>");
        }

        // Map Layers
        if (mapConfig.LayerSettings != null)
        {
            var layerInfo = new List<string>();

            if (mapConfig.LayerSettings.ShowCountryBorders.HasValue)
                layerInfo.Add($"Country Borders: {(mapConfig.LayerSettings.ShowCountryBorders.Value ? "✓" : "✗")}");
            if (mapConfig.LayerSettings.ShowStateBorders.HasValue)
                layerInfo.Add($"State/Province Borders: {(mapConfig.LayerSettings.ShowStateBorders.Value ? "✓" : "✗")}");
            if (mapConfig.LayerSettings.ShowCountyBorders.HasValue)
                layerInfo.Add($"County Borders: {(mapConfig.LayerSettings.ShowCountyBorders.Value ? "✓" : "✗")}");
            if (mapConfig.LayerSettings.ShowCoastlines.HasValue)
                layerInfo.Add($"Coastlines: {(mapConfig.LayerSettings.ShowCoastlines.Value ? "✓" : "✗")}");
            if (mapConfig.LayerSettings.ShowCityNames.HasValue)
                layerInfo.Add($"City Names: {(mapConfig.LayerSettings.ShowCityNames.Value ? "✓" : "✗")}");
            if (mapConfig.LayerSettings.ShowStreets.HasValue)
                layerInfo.Add($"Streets: {(mapConfig.LayerSettings.ShowStreets.Value ? "✓" : "✗")}");
            if (mapConfig.LayerSettings.ShowWaterFeatures.HasValue)
                layerInfo.Add($"Water Features: {(mapConfig.LayerSettings.ShowWaterFeatures.Value ? "✓" : "✗")}");

            if (layerInfo.Any())
            {
                _html.AppendLine("                <h5>Map Layers</h5>");
                _html.AppendLine("                <ul class=\"map-layers-list\">");
                foreach (var layer in layerInfo)
                {
                    _html.AppendLine($"                    <li>{EscapeHtml(layer)}</li>");
                }
                _html.AppendLine("                </ul>");
            }
        }

        _html.AppendLine("            </div>");
    }

    /// <summary>
    /// Append map configuration information for dashboard visuals
    /// </summary>
    private void AppendMapConfiguration(Core.Models.MapConfiguration? mapConfig)
    {
        if (mapConfig == null)
            return;

        _html.AppendLine("                <div class=\"map-configuration\">");
        _html.AppendLine("                    <h6>Map Configuration</h6>");

        // Map Source
        if (!string.IsNullOrEmpty(mapConfig.MapSource))
        {
            _html.AppendLine($"                    <p><strong>Map Provider:</strong> {EscapeHtml(mapConfig.MapSource)}</p>");
        }

        // Geographic Fields
        if (mapConfig.GeographicFields.Any())
        {
            _html.AppendLine("                    <h6>Geographic Fields</h6>");
            _html.AppendLine("                    <div class=\"table-wrapper\">");
            _html.AppendLine("                    <table class=\"geographic-fields-table\">");
            _html.AppendLine("                    <thead>");
            _html.AppendLine("                        <tr>");
            _html.AppendLine("                            <th>Field</th>");
            _html.AppendLine("                            <th>Role</th>");
            _html.AppendLine("                            <th>Shelf</th>");
            _html.AppendLine("                        </tr>");
            _html.AppendLine("                    </thead>");
            _html.AppendLine("                    <tbody>");

            foreach (var geoField in mapConfig.GeographicFields)
            {
                _html.AppendLine("                        <tr>");
                _html.AppendLine($"                            <td>{EscapeHtml(geoField.FieldName)}{(geoField.IsGenerated ? " <em>(generated)</em>" : "")}</td>");
                _html.AppendLine($"                            <td>{EscapeHtml(geoField.GeographicRole)}</td>");
                _html.AppendLine($"                            <td>{EscapeHtml(geoField.Shelf)}</td>");
                _html.AppendLine("                        </tr>");
            }

            _html.AppendLine("                    </tbody>");
            _html.AppendLine("                    </table>");
            _html.AppendLine("                    </div>");
        }

        // Geometry Encoding
        if (mapConfig.HasGeometryEncoding)
        {
            _html.AppendLine("                    <p><strong>Geometry Encoding:</strong> ");
            if (!string.IsNullOrEmpty(mapConfig.GeometryColumn))
            {
                _html.AppendLine($"{EscapeHtml(mapConfig.GeometryColumn)}</p>");
            }
            else
            {
                _html.AppendLine("Yes</p>");
            }
        }

        // Map Styling
        var styleInfo = new List<string>();
        if (!string.IsNullOrEmpty(mapConfig.BaseMapStyle))
        {
            styleInfo.Add($"Style: {EscapeHtml(mapConfig.BaseMapStyle)}");
        }
        if (mapConfig.Washout.HasValue)
        {
            var washoutPercent = (int)(mapConfig.Washout.Value * 100);
            styleInfo.Add($"Washout: {washoutPercent}%");
        }
        if (mapConfig.ShowLabels.HasValue)
        {
            styleInfo.Add($"Labels: {(mapConfig.ShowLabels.Value ? "Shown" : "Hidden")}");
        }

        if (styleInfo.Any())
        {
            _html.AppendLine($"                    <p><strong>Map Style:</strong> {string.Join(", ", styleInfo)}</p>");
        }

        // Map Layers
        if (mapConfig.LayerSettings != null)
        {
            var layerInfo = new List<string>();

            if (mapConfig.LayerSettings.ShowCountryBorders.HasValue)
                layerInfo.Add($"Country Borders: {(mapConfig.LayerSettings.ShowCountryBorders.Value ? "✓" : "✗")}");
            if (mapConfig.LayerSettings.ShowStateBorders.HasValue)
                layerInfo.Add($"State/Province Borders: {(mapConfig.LayerSettings.ShowStateBorders.Value ? "✓" : "✗")}");
            if (mapConfig.LayerSettings.ShowCountyBorders.HasValue)
                layerInfo.Add($"County Borders: {(mapConfig.LayerSettings.ShowCountyBorders.Value ? "✓" : "✗")}");
            if (mapConfig.LayerSettings.ShowCoastlines.HasValue)
                layerInfo.Add($"Coastlines: {(mapConfig.LayerSettings.ShowCoastlines.Value ? "✓" : "✗")}");
            if (mapConfig.LayerSettings.ShowCityNames.HasValue)
                layerInfo.Add($"City Names: {(mapConfig.LayerSettings.ShowCityNames.Value ? "✓" : "✗")}");
            if (mapConfig.LayerSettings.ShowStreets.HasValue)
                layerInfo.Add($"Streets: {(mapConfig.LayerSettings.ShowStreets.Value ? "✓" : "✗")}");
            if (mapConfig.LayerSettings.ShowWaterFeatures.HasValue)
                layerInfo.Add($"Water Features: {(mapConfig.LayerSettings.ShowWaterFeatures.Value ? "✓" : "✗")}");

            if (layerInfo.Any())
            {
                _html.AppendLine("                    <h6>Map Layers</h6>");
                _html.AppendLine("                    <ul class=\"map-layers-list\">");
                foreach (var layer in layerInfo)
                {
                    _html.AppendLine($"                        <li>{EscapeHtml(layer)}</li>");
                }
                _html.AppendLine("                    </ul>");
            }
        }

        _html.AppendLine("                </div>");
    }

    /// <summary>
    /// Append table configuration section for worksheet display
    /// </summary>
    private void AppendTableConfigurationForWorksheet(Core.Models.TableConfiguration? tableConfig, string? title)
    {
        if (tableConfig == null)
            return;

        _html.AppendLine("            <div class=\"table-configuration\">");
        _html.AppendLine("                <h4>Table Configuration</h4>");

        // Display title if present and different from worksheet name
        if (!string.IsNullOrWhiteSpace(title))
        {
            _html.AppendLine($"                <p><strong>Title:</strong> {EscapeHtml(title)}</p>");
        }

        _html.AppendLine($"                <p><strong>Table Type:</strong> {EscapeHtml(tableConfig.TableType)}</p>");

        // Columns table (includes row dimensions as first columns)
        if (tableConfig.Columns.Any())
        {
            _html.AppendLine("                <h6>Columns</h6>");
            _html.AppendLine("                <div class=\"table-wrapper\">");
            _html.AppendLine("                <table class=\"table-columns-table\">");
            _html.AppendLine("                    <thead>");
            _html.AppendLine("                        <tr>");
            _html.AppendLine("                            <th>Order</th>");
            _html.AppendLine("                            <th>Column Name</th>");
            _html.AppendLine("                            <th>Aggregation</th>");
            _html.AppendLine("                            <th>Format</th>");
            _html.AppendLine("                        </tr>");
            _html.AppendLine("                    </thead>");
            _html.AppendLine("                    <tbody>");

            foreach (var column in tableConfig.Columns.OrderBy(c => c.DisplayOrder))
            {
                var friendlyFormat = GetFriendlyFormatName(column.NumberFormat) ?? "-";
                var aggregationHtml = GetAggregationCellHtml(column);
                _html.AppendLine("                        <tr>");
                _html.AppendLine($"                            <td>{column.DisplayOrder + 1}</td>");
                _html.AppendLine($"                            <td>{EscapeHtml(column.DisplayName)}</td>");
                _html.AppendLine($"                            <td>{aggregationHtml}</td>");
                _html.AppendLine($"                            <td>{EscapeHtml(friendlyFormat)}</td>");
                _html.AppendLine("                        </tr>");
            }

            _html.AppendLine("                    </tbody>");
            _html.AppendLine("                </table>");
            _html.AppendLine("                </div>");
        }

        // Table formatting
        if (tableConfig.Formatting != null)
        {
            _html.AppendLine("                <h6>Table Formatting</h6>");
            _html.AppendLine("                <div class=\"table-formatting-grid\">");

            if (!string.IsNullOrEmpty(tableConfig.Formatting.HeaderBackgroundColor))
            {
                _html.AppendLine("                    <div class=\"table-formatting-item\">");
                _html.Append($"                        <strong>Header Background:</strong> {EscapeHtml(tableConfig.Formatting.HeaderBackgroundColor)}");
                _html.AppendLine($" <span class=\"color-swatch\" style=\"background-color: {tableConfig.Formatting.HeaderBackgroundColor};\"></span>");
                _html.AppendLine("                    </div>");
            }

            if (!string.IsNullOrEmpty(tableConfig.Formatting.HeaderTextColor))
            {
                _html.AppendLine("                    <div class=\"table-formatting-item\">");
                _html.Append($"                        <strong>Header Text:</strong> {EscapeHtml(tableConfig.Formatting.HeaderTextColor)}");
                _html.AppendLine($" <span class=\"color-swatch\" style=\"background-color: {tableConfig.Formatting.HeaderTextColor};\"></span>");
                _html.AppendLine("                    </div>");
            }

            if (!string.IsNullOrEmpty(tableConfig.Formatting.RowBandingColor))
            {
                _html.AppendLine("                    <div class=\"table-formatting-item\">");
                _html.Append($"                        <strong>Row Banding:</strong> {EscapeHtml(tableConfig.Formatting.RowBandingColor)}");
                _html.AppendLine($" <span class=\"color-swatch\" style=\"background-color: {tableConfig.Formatting.RowBandingColor};\"></span>");
                _html.AppendLine("                    </div>");
            }

            if (!string.IsNullOrEmpty(tableConfig.Formatting.FontFamily))
            {
                _html.AppendLine("                    <div class=\"table-formatting-item\">");
                _html.AppendLine($"                        <strong>Font:</strong> {EscapeHtml(tableConfig.Formatting.FontFamily)}");
                _html.AppendLine("                    </div>");
            }

            if (tableConfig.Formatting.FontSize.HasValue)
            {
                _html.AppendLine("                    <div class=\"table-formatting-item\">");
                _html.AppendLine($"                        <strong>Font Size:</strong> {tableConfig.Formatting.FontSize}pt");
                _html.AppendLine("                    </div>");
            }

            if (!string.IsNullOrEmpty(tableConfig.Formatting.RowHeaderAlignment))
            {
                _html.AppendLine("                    <div class=\"table-formatting-item\">");
                _html.AppendLine($"                        <strong>Row Alignment:</strong> {EscapeHtml(tableConfig.Formatting.RowHeaderAlignment)}");
                _html.AppendLine("                    </div>");
            }

            _html.AppendLine("                </div>");
        }

        // Banding info
        if (tableConfig.HasRowBanding || tableConfig.HasColumnBanding)
        {
            var bandingInfo = new List<string>();
            if (tableConfig.HasRowBanding) bandingInfo.Add("Row Banding");
            if (tableConfig.HasColumnBanding) bandingInfo.Add("Column Banding");

            _html.AppendLine($"                <p><strong>Banding:</strong> {string.Join(", ", bandingInfo)}</p>");
        }

        _html.AppendLine("            </div>");
    }

    /// <summary>
    /// Append table configuration section for dashboard visual display
    /// </summary>
    private void AppendTableConfiguration(Core.Models.TableConfiguration? tableConfig, string? title)
    {
        if (tableConfig == null)
            return;

        _html.AppendLine("                <div class=\"table-configuration\">");
        _html.AppendLine("                    <h6>Table Configuration</h6>");

        // Display title if present and different from worksheet name
        if (!string.IsNullOrWhiteSpace(title))
        {
            _html.AppendLine($"                    <p><strong>Title:</strong> {EscapeHtml(title)}</p>");
        }

        _html.AppendLine($"                    <p><strong>Table Type:</strong> {EscapeHtml(tableConfig.TableType)}</p>");

        // Columns table (includes row dimensions as first columns)
        if (tableConfig.Columns.Any())
        {
            _html.AppendLine("                    <h6>Columns</h6>");
            _html.AppendLine("                    <div class=\"table-wrapper\">");
            _html.AppendLine("                    <table class=\"table-columns-table\">");
            _html.AppendLine("                        <thead>");
            _html.AppendLine("                            <tr>");
            _html.AppendLine("                                <th>Order</th>");
            _html.AppendLine("                                <th>Column Name</th>");
            _html.AppendLine("                                <th>Aggregation</th>");
            _html.AppendLine("                                <th>Format</th>");
            _html.AppendLine("                            </tr>");
            _html.AppendLine("                        </thead>");
            _html.AppendLine("                        <tbody>");

            foreach (var column in tableConfig.Columns.OrderBy(c => c.DisplayOrder))
            {
                var friendlyFormat = GetFriendlyFormatName(column.NumberFormat) ?? "-";
                var aggregationHtml = GetAggregationCellHtml(column);
                _html.AppendLine("                            <tr>");
                _html.AppendLine($"                                <td>{column.DisplayOrder + 1}</td>");
                _html.AppendLine($"                                <td>{EscapeHtml(column.DisplayName)}</td>");
                _html.AppendLine($"                                <td>{aggregationHtml}</td>");
                _html.AppendLine($"                                <td>{EscapeHtml(friendlyFormat)}</td>");
                _html.AppendLine("                            </tr>");
            }

            _html.AppendLine("                        </tbody>");
            _html.AppendLine("                    </table>");
            _html.AppendLine("                    </div>");
        }

        // Table formatting
        if (tableConfig.Formatting != null)
        {
            _html.AppendLine("                    <h6>Table Formatting</h6>");
            _html.AppendLine("                    <div class=\"table-formatting-grid\">");

            if (!string.IsNullOrEmpty(tableConfig.Formatting.HeaderBackgroundColor))
            {
                _html.AppendLine("                        <div class=\"table-formatting-item\">");
                _html.Append($"                            <strong>Header Background:</strong> {EscapeHtml(tableConfig.Formatting.HeaderBackgroundColor)}");
                _html.AppendLine($" <span class=\"color-swatch\" style=\"background-color: {tableConfig.Formatting.HeaderBackgroundColor};\"></span>");
                _html.AppendLine("                        </div>");
            }

            if (!string.IsNullOrEmpty(tableConfig.Formatting.HeaderTextColor))
            {
                _html.AppendLine("                        <div class=\"table-formatting-item\">");
                _html.Append($"                            <strong>Header Text:</strong> {EscapeHtml(tableConfig.Formatting.HeaderTextColor)}");
                _html.AppendLine($" <span class=\"color-swatch\" style=\"background-color: {tableConfig.Formatting.HeaderTextColor};\"></span>");
                _html.AppendLine("                        </div>");
            }

            if (!string.IsNullOrEmpty(tableConfig.Formatting.RowBandingColor))
            {
                _html.AppendLine("                        <div class=\"table-formatting-item\">");
                _html.Append($"                            <strong>Row Banding:</strong> {EscapeHtml(tableConfig.Formatting.RowBandingColor)}");
                _html.AppendLine($" <span class=\"color-swatch\" style=\"background-color: {tableConfig.Formatting.RowBandingColor};\"></span>");
                _html.AppendLine("                        </div>");
            }

            if (!string.IsNullOrEmpty(tableConfig.Formatting.FontFamily))
            {
                _html.AppendLine("                        <div class=\"table-formatting-item\">");
                _html.AppendLine($"                            <strong>Font:</strong> {EscapeHtml(tableConfig.Formatting.FontFamily)}");
                _html.AppendLine("                        </div>");
            }

            if (tableConfig.Formatting.FontSize.HasValue)
            {
                _html.AppendLine("                        <div class=\"table-formatting-item\">");
                _html.AppendLine($"                            <strong>Font Size:</strong> {tableConfig.Formatting.FontSize}pt");
                _html.AppendLine("                        </div>");
            }

            if (!string.IsNullOrEmpty(tableConfig.Formatting.RowHeaderAlignment))
            {
                _html.AppendLine("                        <div class=\"table-formatting-item\">");
                _html.AppendLine($"                            <strong>Row Alignment:</strong> {EscapeHtml(tableConfig.Formatting.RowHeaderAlignment)}");
                _html.AppendLine("                        </div>");
            }

            _html.AppendLine("                    </div>");
        }

        // Banding info
        if (tableConfig.HasRowBanding || tableConfig.HasColumnBanding)
        {
            var bandingInfo = new List<string>();
            if (tableConfig.HasRowBanding) bandingInfo.Add("Row Banding");
            if (tableConfig.HasColumnBanding) bandingInfo.Add("Column Banding");

            _html.AppendLine($"                    <p><strong>Banding:</strong> {string.Join(", ", bandingInfo)}</p>");
        }

        _html.AppendLine("                </div>");
    }

    /// <summary>
    /// Convert technical format string to friendly format name
    /// </summary>
    private string? GetFriendlyFormatName(string? formatString)
    {
        if (string.IsNullOrEmpty(formatString))
            return null;

        // Currency formats
        if (formatString.Contains("$") || formatString.Contains("c\""))
            return "Currency";

        // Percentage formats
        if (formatString.Contains("%") || formatString.StartsWith("p"))
            return "Percentage";

        // Number formats with thousands separator
        if (formatString.Contains("#,##0"))
            return "Number";

        // Date formats
        if (formatString.Contains("yyyy") || formatString.Contains("MM") || formatString.Contains("dd"))
            return "Date";

        // Default: return the original format string
        return formatString;
    }

    private string EscapeHtml(string text)
    {
        if (string.IsNullOrEmpty(text))
            return string.Empty;

        return text
            .Replace("&", "&amp;")
            .Replace("<", "&lt;")
            .Replace(">", "&gt;")
            .Replace("\"", "&quot;")
            .Replace("'", "&#39;");
    }

    /// <summary>
    /// Format a field name with explanatory annotation for known Tableau patterns
    /// </summary>
    private string FormatFieldWithAnnotation(string fieldName)
    {
        if (string.IsNullOrEmpty(fieldName))
            return EscapeHtml(fieldName);

        var annotation = GetFieldAnnotation(fieldName);
        if (!string.IsNullOrEmpty(annotation))
        {
            return $"{EscapeHtml(fieldName)} <span class=\"field-annotation\" title=\"{EscapeHtml(annotation)}\">[?]</span>";
        }

        return EscapeHtml(fieldName);
    }

    /// <summary>
    /// Get explanatory annotation for known Tableau calculation patterns
    /// </summary>
    private string? GetFieldAnnotation(string fieldName)
    {
        if (string.IsNullOrEmpty(fieldName))
            return null;

        var normalizedName = fieldName.Trim().ToLowerInvariant();

        // Avg(0), Min(0), Max(0), Sum(0) - placeholder/reference line patterns
        if (normalizedName == "avg(0)" || normalizedName == "min(0)" ||
            normalizedName == "max(0)" || normalizedName == "sum(0)")
        {
            return "Placeholder field that always returns 0. Commonly used for reference lines, axis alignment, or positioning elements in dual-axis charts.";
        }

        // Rounded Corners - common Tableau trick
        if (normalizedName == "rounded corners" || normalizedName.Contains("rounded corner"))
        {
            return "Design element used to create rounded corner effects in dashboard visualizations.";
        }

        // Tableau internal field notation: aggregation:fieldname:type
        // Examples: sum:X:qk, attr:Name:nk, count:Number:qk
        if (System.Text.RegularExpressions.Regex.IsMatch(fieldName, @"^[a-z]+:[^:]+:[a-z]+$", System.Text.RegularExpressions.RegexOptions.IgnoreCase))
        {
            var parts = fieldName.Split(':');
            if (parts.Length == 3)
            {
                var aggr = parts[0].ToLowerInvariant();
                var field = parts[1];
                var type = parts[2].ToLowerInvariant();

                var aggrDesc = aggr switch
                {
                    "sum" => "Sum",
                    "avg" => "Average",
                    "min" => "Minimum",
                    "max" => "Maximum",
                    "count" => "Count",
                    "countd" => "Count Distinct",
                    "attr" => "Attribute",
                    "none" => "No aggregation",
                    _ => aggr.ToUpper()
                };

                var typeDesc = type switch
                {
                    "qk" => "quantitative",
                    "nk" => "nominal/categorical",
                    "ok" => "ordinal",
                    _ => type
                };

                return $"Tableau internal notation: {aggrDesc} of '{field}' ({typeDesc} field)";
            }
        }

        return null;
    }

    private string GetConnectionTypeDescription(string connectionType)
    {
        return connectionType?.ToLower() switch
        {
            "federated" => "Combines data from multiple sources (e.g., Excel + database) into a single data model using joins or unions.",
            "excel" or "excel-direct" => "Direct connection to a Microsoft Excel workbook.",
            "textscan" => "Connection to a text file (CSV, TSV, or delimited text).",
            "sqlserver" => "Connection to Microsoft SQL Server database.",
            "postgres" => "Connection to PostgreSQL database.",
            "mysql" => "Connection to MySQL database.",
            "oracle" => "Connection to Oracle database.",
            "snowflake" => "Connection to Snowflake cloud data warehouse.",
            "bigquery" => "Connection to Google BigQuery.",
            "redshift" => "Connection to Amazon Redshift.",
            "databricks" => "Connection to Databricks.",
            "extract" => "Uses a Tableau data extract (.hyper file) for optimized performance.",
            "live" => "Real-time connection that queries the source directly.",
            _ => string.Empty
        };
    }

    /// <summary>
    /// Generate HTML for aggregation cell with clickable link for Custom aggregations
    /// </summary>
    private string GetAggregationCellHtml(TableColumn column)
    {
        if (column.Aggregation == "Custom")
        {
            // Try to find the field in any datasource to generate correct link
            string? anchorLink = null;
            if (!string.IsNullOrEmpty(column.CalculationName))
            {
                foreach (var dataSource in _workbook.DataSources ?? Enumerable.Empty<Core.Models.DataSource>())
                {
                    var field = dataSource.Fields?.FirstOrDefault(f => f.Name == column.CalculationName);
                    if (field != null)
                    {
                        var dsId = $"ds-{ToHtmlId(dataSource.Name)}-field-{ToHtmlId(column.CalculationName)}";
                        anchorLink = $"#{dsId}";
                        break;
                    }
                }
            }

            if (anchorLink != null)
            {
                return $"<a href=\"{anchorLink}\" class=\"calc-link\" title=\"View calculation details\">{EscapeHtml(column.Aggregation)} ↗</a>";
            }

            // Fallback: calculation not found in datasources, show with tooltip explanation
            return $"<span class=\"calc-no-link\" title=\"This custom calculation is not available in the Data Sources section (may be a worksheet-specific calculation or inline formula)\">{EscapeHtml(column.Aggregation)}</span>";
        }

        return EscapeHtml(column.Aggregation);
    }


    private string GetUserFriendlyControlType(Core.Models.Filter filter)
    {
        // For parameters, translate "Parameter Control" to a user-friendly control type
        if (filter.Type == "parameter" || filter.ControlType == "Parameter Control")
        {
            // Parameters are typically single-select dropdowns from a user perspective
            return "Single-select dropdown";
        }

        return filter.ControlType ?? string.Empty;
    }

    private string? GetFieldLink(string fieldName, string? dataSourceName)
    {
        if (string.IsNullOrEmpty(fieldName) || string.IsNullOrEmpty(dataSourceName))
            return null;

        // Find the data source
        var dataSource = _workbook.DataSources?.FirstOrDefault(ds =>
            (ds.Name ?? ds.Caption) == dataSourceName);

        if (dataSource == null)
            return null;

        // Check if the field exists in this data source
        var field = dataSource.Fields?.FirstOrDefault(f => f.Name == fieldName);

        if (field != null)
        {
            // Create anchor link to the data source section
            var dsId = $"ds-{ToHtmlId(dataSourceName)}-field-{ToHtmlId(fieldName)}";
            return $"#{dsId}";
        }

        return null;
    }

    private string GetDefaultSelection(Core.Models.Filter filter)
    {
        // If there's an explicit default selection, use it
        if (!string.IsNullOrEmpty(filter.DefaultSelection))
        {
            return filter.DefaultSelection;
        }

        // For multi-select filters (dropdown, list), default is typically "All"
        if (filter.ControlType == "Multi-select dropdown" ||
            filter.ControlType == "List" ||
            filter.ControlType == "Compact list" ||
            filter.ControlType == "checkdropdown")
        {
            return "All";
        }

        return string.Empty;
    }
}
