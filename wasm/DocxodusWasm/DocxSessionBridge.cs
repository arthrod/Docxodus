using System;
using System.Collections.Concurrent;
using System.Linq;
using System.Runtime.InteropServices.JavaScript;
using System.Runtime.Versioning;
using System.Text;
using System.Text.Json;
using Docxodus;

namespace DocxodusWasm;

/// <summary>
/// JSExport bridge for <see cref="DocxSession"/>. Sessions live on the .NET heap
/// and persist across JSExport calls — keyed by an integer handle returned from
/// <see cref="OpenSession"/>. JS-side code must call <see cref="CloseSession"/>
/// when done; sessions are not eligible for GC otherwise.
/// </summary>
[SupportedOSPlatform("browser")]
public static partial class DocxSessionBridge
{
    private static readonly ConcurrentDictionary<int, DocxSession> _sessions = new();
    private static int _nextId;

    [JSExport]
    public static int OpenSession(byte[] bytes, string settingsJson)
    {
        var settings = ParseSettings(settingsJson);
        var session = new DocxSession(bytes, settings);
        var id = System.Threading.Interlocked.Increment(ref _nextId);
        _sessions[id] = session;
        return id;
    }

    [JSExport]
    public static void CloseSession(int handle)
    {
        if (_sessions.TryRemove(handle, out var s)) s.Dispose();
    }

    [JSExport]
    public static string Project(int handle) => SerializeProjection(Get(handle).Project());

    [JSExport]
    public static string ReplaceText(int h, string anchor, string md) =>
        Serialize(Get(h).ReplaceText(anchor, md));

    [JSExport]
    public static string DeleteBlock(int h, string anchor) =>
        Serialize(Get(h).DeleteBlock(anchor));

    [JSExport]
    public static string InsertParagraph(int h, string anchor, string posStr, string md) =>
        Serialize(Get(h).InsertParagraph(anchor, ParsePos(posStr), md));

    [JSExport]
    public static string SplitParagraph(int h, string anchor, int offset) =>
        Serialize(Get(h).SplitParagraph(anchor, offset));

    [JSExport]
    public static string MergeParagraphs(int h, string first, string second) =>
        Serialize(Get(h).MergeParagraphs(first, second));

    [JSExport]
    public static string ApplyFormat(int h, string anchor, string spanJson, string opJson)
    {
        CharSpan? span = null;
        if (!string.IsNullOrEmpty(spanJson))
        {
            using var doc = JsonDocument.Parse(spanJson);
            span = new CharSpan(
                doc.RootElement.GetProperty("start").GetInt32(),
                doc.RootElement.GetProperty("length").GetInt32());
        }
        var op = ParseFormatOp(opJson);
        return Serialize(Get(h).ApplyFormat(anchor, span, op));
    }

    [JSExport]
    public static string SetParagraphStyle(int h, string anchor, string styleId) =>
        Serialize(Get(h).SetParagraphStyle(anchor, styleId));

    [JSExport]
    public static string SetListLevel(int h, string anchor, int delta) =>
        Serialize(Get(h).SetListLevel(anchor, delta));

    [JSExport]
    public static string RemoveListMembership(int h, string anchor) =>
        Serialize(Get(h).RemoveListMembership(anchor));

    [JSExport]
    public static string ReplaceCellContent(int h, string anchor, string md) =>
        Serialize(Get(h).ReplaceCellContent(anchor, md));

    [JSExport]
    public static string RawGetXml(int h, string anchor) => Get(h).Raw.GetXml(anchor);

    [JSExport]
    public static string RawInsertXml(int h, string anchor, string posStr, string xml) =>
        Serialize(Get(h).Raw.InsertXml(anchor, ParsePos(posStr), xml));

    [JSExport]
    public static string RawReplaceXml(int h, string anchor, string xml) =>
        Serialize(Get(h).Raw.ReplaceXml(anchor, xml));

    [JSExport]
    public static bool Undo(int h) => Get(h).Undo();

    [JSExport]
    public static bool Redo(int h) => Get(h).Redo();

    [JSExport]
    public static byte[] Save(int h) => Get(h).Save();

    // ─── Helpers ────────────────────────────────────────────────────────

    private static DocxSession Get(int handle)
    {
        if (!_sessions.TryGetValue(handle, out var s))
            throw new ArgumentException($"unknown session handle: {handle}");
        return s;
    }

    private static Position ParsePos(string s) =>
        string.Equals(s, "before", StringComparison.OrdinalIgnoreCase) ? Position.Before : Position.After;

    private static DocxSessionSettings ParseSettings(string settingsJson)
    {
        if (string.IsNullOrEmpty(settingsJson)) return new DocxSessionSettings();
        using var doc = JsonDocument.Parse(settingsJson);
        var root = doc.RootElement;
        int undoDepth = TryGetInt(root, "undoDepth", 50);
        bool validateRawOps = TryGetBool(root, "validateRawOps", false);
        var trackedStr = TryGetString(root, "trackedChanges", "accept");
        var tracked = trackedStr switch
        {
            "render_inline" => TrackedChangeMode.RenderInline,
            "strip_deletions" => TrackedChangeMode.StripDeletions,
            _ => TrackedChangeMode.Accept,
        };
        var revisionAuthor = TryGetString(root, "revisionAuthor", null);
        bool persistAnchorIds = TryGetBool(root, "persistAnchorIds", false);
        return new DocxSessionSettings
        {
            UndoDepth = undoDepth,
            ValidateRawOps = validateRawOps,
            TrackedChanges = tracked,
            RevisionAuthor = revisionAuthor,
            PersistAnchorIds = persistAnchorIds,
        };
    }

    private static FormatOp ParseFormatOp(string json)
    {
        if (string.IsNullOrEmpty(json)) return new FormatOp();
        using var doc = JsonDocument.Parse(json);
        var root = doc.RootElement;
        return new FormatOp
        {
            Bold = TryGetBoolNullable(root, "bold"),
            Italic = TryGetBoolNullable(root, "italic"),
            Underline = TryGetBoolNullable(root, "underline"),
            Strike = TryGetBoolNullable(root, "strike"),
            Code = TryGetBoolNullable(root, "code"),
            Color = TryGetString(root, "color", null),
            RunStyle = TryGetString(root, "runStyle", null),
        };
    }

    private static int TryGetInt(JsonElement root, string name, int fallback) =>
        root.TryGetProperty(name, out var v) && v.ValueKind == JsonValueKind.Number ? v.GetInt32() : fallback;
    private static bool TryGetBool(JsonElement root, string name, bool fallback) =>
        root.TryGetProperty(name, out var v) && (v.ValueKind == JsonValueKind.True || v.ValueKind == JsonValueKind.False) ? v.GetBoolean() : fallback;
    private static bool? TryGetBoolNullable(JsonElement root, string name) =>
        root.TryGetProperty(name, out var v) && (v.ValueKind == JsonValueKind.True || v.ValueKind == JsonValueKind.False) ? v.GetBoolean() : (bool?)null;
    private static string? TryGetString(JsonElement root, string name, string? fallback) =>
        root.TryGetProperty(name, out var v) && v.ValueKind == JsonValueKind.String ? v.GetString() : fallback;

    private static string Serialize(EditResult r)
    {
        var sb = new StringBuilder(256);
        sb.Append("{\"success\":").Append(r.Success ? "true" : "false");
        if (r.Error is not null)
        {
            sb.Append(",\"error\":{")
              .Append("\"code\":\"").Append(EnumToSnake(r.Error.Code)).Append('"')
              .Append(",\"message\":").Append(JsonString(r.Error.Message));
            if (r.Error.AnchorId is not null)
                sb.Append(",\"anchorId\":").Append(JsonString(r.Error.AnchorId));
            sb.Append('}');
        }
        sb.Append(",\"created\":"); AppendAnchorArray(sb, r.Created);
        sb.Append(",\"removed\":"); AppendAnchorArray(sb, r.Removed);
        sb.Append(",\"modified\":"); AppendAnchorArray(sb, r.Modified);
        if (r.Patch is not null)
        {
            sb.Append(",\"patch\":{")
              .Append("\"scopeAnchorId\":").Append(JsonString(r.Patch.ScopeAnchorId))
              .Append(",\"markdown\":").Append(JsonString(r.Patch.Markdown))
              .Append('}');
        }
        sb.Append('}');
        return sb.ToString();
    }

    private static void AppendAnchorArray(StringBuilder sb, System.Collections.Generic.IReadOnlyList<Anchor> anchors)
    {
        sb.Append('[');
        for (int i = 0; i < anchors.Count; i++)
        {
            if (i > 0) sb.Append(',');
            var a = anchors[i];
            sb.Append('{')
              .Append("\"id\":").Append(JsonString(a.Id))
              .Append(",\"kind\":").Append(JsonString(a.Kind))
              .Append(",\"scope\":").Append(JsonString(a.Scope))
              .Append(",\"unid\":").Append(JsonString(a.Unid))
              .Append('}');
        }
        sb.Append(']');
    }

    private static string SerializeProjection(MarkdownProjection p)
    {
        var sb = new StringBuilder(p.Markdown.Length + 200);
        sb.Append("{\"markdown\":").Append(JsonString(p.Markdown));
        sb.Append(",\"anchorIndex\":{");
        bool first = true;
        foreach (var kv in p.AnchorIndex)
        {
            if (!first) sb.Append(',');
            first = false;
            sb.Append(JsonString(kv.Key)).Append(":{")
              .Append("\"partUri\":").Append(JsonString(kv.Value.PartUri))
              .Append(",\"unid\":").Append(JsonString(kv.Value.Unid))
              .Append(",\"kind\":").Append(JsonString(kv.Value.Anchor.Kind))
              .Append(",\"scope\":").Append(JsonString(kv.Value.Anchor.Scope))
              .Append('}');
        }
        sb.Append("}}");
        return sb.ToString();
    }

    private static string JsonString(string s)
    {
        var sb = new StringBuilder(s.Length + 2);
        sb.Append('"');
        foreach (var c in s)
        {
            switch (c)
            {
                case '"': sb.Append("\\\""); break;
                case '\\': sb.Append("\\\\"); break;
                case '\n': sb.Append("\\n"); break;
                case '\r': sb.Append("\\r"); break;
                case '\t': sb.Append("\\t"); break;
                case '\b': sb.Append("\\b"); break;
                case '\f': sb.Append("\\f"); break;
                default:
                    if (c < 0x20) sb.Append("\\u").Append(((int)c).ToString("X4"));
                    else sb.Append(c);
                    break;
            }
        }
        sb.Append('"');
        return sb.ToString();
    }

    private static string EnumToSnake(EditErrorCode code)
    {
        var s = code.ToString();
        var sb = new StringBuilder(s.Length + 4);
        for (int i = 0; i < s.Length; i++)
        {
            if (i > 0 && char.IsUpper(s[i])) sb.Append('_');
            sb.Append(char.ToLowerInvariant(s[i]));
        }
        return sb.ToString();
    }
}
