# LibreOffice adapter

Wraps LibreOffice Writer behind adapter protocol v1. This adapter targets the
Ubuntu CI environment, where headless `soffice` and pyuno work normally. It
does not try to support macOS local execution: pyuno is blocked there by
Launch Constraints, which is why the upstream oracle recipe uses Basic macro
injection instead.

**What is library API vs glue:** `acceptAllTrackedChanges` and
`rejectAllTrackedChanges` use LibreOffice's own UNO dispatch commands,
`.uno:AcceptAllTrackedChanges` and `.uno:RejectAllTrackedChanges`.
`replaceFirstTextOccurrence` uses the document's UNO search API
(`createSearchDescriptor`, `findFirst`, and `setString`) with case-sensitive,
non-regex search. The adapter glue is process lifecycle, a throwaway
`UserInstallation` profile, protocol parsing, hidden document load/save, and
DOCX export through the `MS Word 2007 XML` filter.

**Paragraph-local replacement semantics:** the scenario DSL defines
`replaceFirstTextOccurrence` as the first paragraph in document order whose
text contains `findText`; matches spanning paragraph boundaries are out of
scope. LibreOffice `findFirst` searches text ranges in document order, and
with non-regex, case-sensitive search it returns the first in-document match
rather than an adapter-computed rewrite. Since cross-paragraph matches are out
of scope, that first match is in the first matching paragraph and honors the
DSL's paragraph-local-first rule.

**Known upstream caveat:** LibreOffice drops `w:ins` provenance on save when
an insertion's entire content has been deleted (reproduced on LO 25.8.7.3).
Scenarios touching nested del-in-ins would show honest `fail` cells — known
upstream behavior, not suite breakage.

**Maintenance policy:** best-effort. CI installs LibreOffice from Ubuntu apt;
version drift is intentionally visible, and the weekly cron re-runs the matrix
so upstream breakage surfaces as `error` cells rather than silent rot.
