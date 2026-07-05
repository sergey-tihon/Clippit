using Clippit.Cli.Infrastructure;

namespace Clippit.Cli.Commands.Word.SimplifyMarkup;

internal sealed record WordSimplifyMarkupOptions(
    string InputPath,
    string? OutputPath,
    bool Force,
    bool All,
    bool AcceptRevisions,
    bool RemoveRsidInfo,
    bool RemoveMarkupForDocComp,
    bool RemoveComments,
    bool RemoveBookmarks,
    bool RemoveContentControls,
    bool RemoveEndAndFootnotes,
    bool RemoveFieldCodes,
    bool RemoveGoBackBookmark,
    bool RemoveHyperlinks,
    bool RemoveLastRenderedPageBreak,
    bool RemovePermissions,
    bool RemoveProof,
    bool RemoveSmartTags,
    bool RemoveSoftHyphens,
    bool RemoveWebHidden,
    bool ReplaceTabsWithSpaces,
    bool NormalizeXml,
    OutputFormat Format,
    bool Quiet
);
