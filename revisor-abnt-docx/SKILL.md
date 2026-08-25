---
name: revisor-abnt-docx
description: Review and format an existing Brazilian academic DOCX/TCC according to ABNT NBR 14724:2024, NBR 6023 (2018 plus public 2025 updates), NBR 10520:2023, NBR 6028:2021, NBR 6027:2012, NBR 6024:2012, and NBR 6022:2018, preserving content and returning a separate formatted copy. Use for DOCX TCC, monograph, dissertation, thesis, or academic article formatting; not for inventing citations or bibliographic data.
---

# Revisor ABNT DOCX

Use this skill when the user provides an existing Word document and asks for an ABNT formatting review. The result is a new DOCX; never overwrite the source file.

## Normative scope and precedence

- Use the current editions supplied or explicitly required by the user. The default profile is NBR 14724:2024 (academic work presentation), NBR 6023:2018 with public 2025 updates (references), NBR 10520:2023 (citations), NBR 6028:2021 (abstracts), NBR 6027:2012 (table of contents), NBR 6024:2012 (progressive section numbering), and NBR 6022:2018 (scientific article presentation).
- The operational baseline comes from [references/normas-abnt.md](references/normas-abnt.md) and the public institutional manuals indexed in [references/fontes-publicas.md](references/fontes-publicas.md). They are secondary support materials, not the paid ABNT standards. The full standard texts were neither copied nor downloaded; when a rule is ambiguous or absent, report it instead of guessing.
- An institutional manual, template, or course rule supplied by the user takes precedence over this baseline. Identify and follow it before applying automated changes.
- Do not claim certification or absolute compliance: a DOCX cannot establish source accuracy, authorship, originality, or every special reference case automatically.
- Never invent, complete, reorder, delete, or silently rewrite bibliographic data, quotations, author names, dates, pages, URLs, or access dates. Flag missing or ambiguous data for the user.

## Workflow

1. Inspect the input DOCX, all supplied institutional instructions, and the existing section/layout structure. Determine whether it has the required pre-textual, textual, and post-textual elements; an abstract/abstract in another language; a table of contents; and a coherent section-numbering hierarchy.
2. Preserve the original. Run `scripts/audit_docx_abnt.py` first and keep its JSON report with the work files.
3. Run `scripts/format_docx_abnt.py INPUT.docx --out OUTPUT_ABNT.docx`. The body formatting starts only when the textual part is identified (normally `INTRODUÇÃO`); cover, folha de rosto, dedicatória and epígrafe therefore keep their own formatting. The formatter applies safe layout normalization, preserves Word automatic numbering, produces an action/issue report and deliberately does not fabricate citations or reference entries. Do not request automatic page-number insertion unless the document already has a confirmed section break before the textual part.
4. Review citations, references, abstracts, table of contents, and section numbering using [references/normas-abnt.md](references/normas-abnt.md) (operational baseline), [references/fontes-publicas.md](references/fontes-publicas.md) (source scope and limits), and [references/abnt-review-rules.md](references/abnt-review-rules.md) (checklist). Correct only information demonstrably supported by the document or user-provided sources. For unresolved cases, add a concise review report instead of guessing.
5. Update the table of contents fields in Word after formatting. Do not type static page numbers or claim that automated rendering updated them.
6. Render the final DOCX to page images with the documents renderer and inspect every page. Fix visible defects, then render again. Retain the original and deliver only the final DOCX unless the user asks for the report too.

## Required output behavior

- File name: preserve the input stem and append `_ABNT.docx` unless the user gives an output name.
- Preserve text, tables, figures, front matter, section breaks, and meaningful italics/bold. Apply global body normalization carefully; do not use a blanket direct-format clearing operation that would erase emphasis in quotations or references.
- Apply the standard textual baseline unless the institutional manual says otherwise: A4, 3 cm top/left margins, 2 cm right/bottom margins, readable 12 pt body font, 1.5 line spacing, justified body paragraphs, and 1.25 cm first-line indentation. Long quotations, notes, captions, and references require their own treatment.
- Do not move page numbering to the first visible page without confirming how the front matter is structured. When used, page numbers belong in the upper-right header and textual pages are counted according to the institution's rule.
- Maintain a separate `*_abnt_report.json` audit report. Clearly distinguish applied formatting from items requiring human/source validation.
- Use real Word heading styles and a consistent progressive-numbering hierarchy. Do not fabricate a sumario with manually typed page numbers. When automatic Word fields are absent, report that the user must update or insert the field-based sumario in Word.
- Preserve the language and substantive content of resumo and abstract. Check their placement, heading, keywords, and visual formatting; do not translate, shorten, expand, or invent keywords without express authorization.

## Citation and reference review

- Direct quotations need author, year, and page or locator when available. Long quotations need the reduced-size, single-spaced, indented presentation specified by the applicable norm/manual, without quotation marks.
- Indirect citations must faithfully correspond to an identifiable source. Do not add page numbers or authors by assumption.
- In the references section, use single spacing within entries, left alignment, no first-line indent, and a blank-line equivalent separation between entries. Keep the sequence and punctuation unless verified source metadata authorizes a correction.
- Match in-text cited authors/years to the reference list and report likely unmatched items. Automated matching is only a lead for manual confirmation.

## Structural review

- Check that the document organization follows NBR 14724 and the institutional manual: pre-textual elements, textual development, and post-textual elements. Report missing expected elements rather than creating personal, institutional, or catalog data.
- Check NBR 6028 aspects for resumo/abstract: heading, one-paragraph presentation when required, keywords, and correspondence between the Portuguese resumo and foreign-language abstract. Treat word-count compliance and translation accuracy as human checks unless the institutional manual defines a determinable rule.
- Check NBR 6027 aspects for the sumario: title, hierarchy, section labels, and page-number field behavior. It must reflect the real heading structure.
- Check NBR 6024 aspects for progressive numbering: levels must be logically nested and use consistent separators. Do not renumber a document when it would alter legally or academically meaningful cross-references without first confirming the intended hierarchy. Some documents number chapter titles with Word's native multilevel-list numbering (a `List Paragraph`/`numPr` paragraph, flagged by the auditor as `list_numbered_heading_candidates`) instead of a Heading style or literal digits; the formatter deliberately does not rewrite list-numbering tab stops (doing so risks breaking the numbering), so visually confirm the number-to-title spacing matches "N TÍTULO" instead of assuming it is centered or misaligned.
- When the deliverable is an article rather than a monograph/TCC, apply NBR 6022 to its presentation: title, author identification when supplied, abstracts, keywords, textual sections, references, and article-specific elements. Do not impose a TCC cover-sheet or pre-textual structure on an article.

## Verification and handoff

- For document edits, use the bundled workspace Python and the documents renderer. Before the first edit operation, run the documents operation marker required by the `documents` skill.
- Visually inspect the rendered pages for margins, paragraph rhythm, page breaks, headings, quotations, tables, figures, headers, and footers. If rendering is unavailable, disclose that visual QA could not be completed.
- Deliver the final DOCX and succinctly list: formatting applied, suspected citation/reference/structure issues, and any remaining manual checks.
