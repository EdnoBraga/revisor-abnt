"""Gera um relatório em PDF, legível por humanos, do que o motor ABNT fez.

O usuário pediu explicitamente correção automática em vez de auditoria, e um
relatório em PDF em vez de JSON: "quero que faça a correção automática no
texto inserido e mostrar o que foi feito em um PDF, não json" / "em todas as
correções necessárias". Este módulo consome o dicionário de relatório que
`abnt_engine.apply_formatting()` já produz (o mesmo usado para o JSON) e o
apresenta em texto corrido, sem exigir que o usuário interprete códigos ou
estrutura de dados.
"""

from __future__ import annotations

from datetime import datetime
from pathlib import Path
from typing import Any

from reportlab.lib import colors
from reportlab.lib.enums import TA_LEFT
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import cm
from reportlab.platypus import (
    ListFlowable,
    ListItem,
    Paragraph,
    SimpleDocTemplate,
    Spacer,
    Table,
    TableStyle,
)

DOCUMENT_TYPE_LABELS = {"tcc": "TCC / monografia / dissertação / tese", "article": "Artigo científico"}
CITATION_SYSTEM_LABELS = {"auto": "Detectado automaticamente", "author-date": "Autor-data", "numeric": "Numérico"}
TOC_MODE_LABELS = {
    "audit": "Preservar e auditar (sem inserir sumário automaticamente)",
    "insert-if-empty": "Criar/inserir campo automático de sumário",
}
PAGINATION_MODE_LABELS = {
    "audit": "Preservar e auditar (sem corrigir numeração automaticamente)",
    "request": "Corrigir numeração automaticamente (NBR 14724)",
}
INSTITUTION_LABELS = {"generic": "Perfil geral ABNT", "cgaem": "CGAEM/ESFCEx (artigo científico)"}
SEVERITY_LABELS = {"error": "Correção necessária", "warning": "Atenção", "info": "Informativo"}
SEVERITY_ORDER = {"error": 0, "warning": 1, "info": 2}


def _styles() -> dict[str, ParagraphStyle]:
    base = getSampleStyleSheet()
    return {
        "title": ParagraphStyle("ReportTitle", parent=base["Title"], fontSize=18, spaceAfter=4, textColor=colors.HexColor("#1a2b4c")),
        "subtitle": ParagraphStyle("ReportSubtitle", parent=base["Normal"], fontSize=10, textColor=colors.HexColor("#555555"), spaceAfter=14),
        "h2": ParagraphStyle("ReportH2", parent=base["Heading2"], fontSize=13, spaceBefore=16, spaceAfter=6, textColor=colors.HexColor("#1a2b4c")),
        "body": ParagraphStyle("ReportBody", parent=base["Normal"], fontSize=10.5, leading=15, alignment=TA_LEFT, spaceAfter=4),
        "item": ParagraphStyle("ReportItem", parent=base["Normal"], fontSize=10.5, leading=15, alignment=TA_LEFT),
        "empty": ParagraphStyle("ReportEmpty", parent=base["Normal"], fontSize=10.5, leading=15, textColor=colors.HexColor("#666666"), fontName="Helvetica-Oblique"),
        "footer": ParagraphStyle("ReportFooter", parent=base["Normal"], fontSize=8.5, leading=12, textColor=colors.HexColor("#777777")),
    }


def _config_table(config: dict[str, Any], styles: dict[str, ParagraphStyle]) -> Table:
    rows = [
        ("Tipo de documento", DOCUMENT_TYPE_LABELS.get(config.get("document_type"), config.get("document_type", "—"))),
        ("Sistema de citação", CITATION_SYSTEM_LABELS.get(config.get("citation_system"), config.get("citation_system", "—"))),
        ("Fonte", config.get("font", "—")),
        ("Sumário", TOC_MODE_LABELS.get(config.get("toc_mode"), config.get("toc_mode", "—"))),
        ("Paginação", PAGINATION_MODE_LABELS.get(config.get("pagination_mode"), config.get("pagination_mode", "—"))),
        ("Perfil institucional", INSTITUTION_LABELS.get(config.get("institution"), config.get("institution", "—"))),
        ("Ordenar referências", "Sim" if config.get("order_references") else "Não"),
    ]
    data = [[Paragraph(f"<b>{label}</b>", styles["item"]), Paragraph(str(value), styles["item"])] for label, value in rows]
    table = Table(data, colWidths=[5 * cm, 10.5 * cm])
    table.setStyle(
        TableStyle(
            [
                ("GRID", (0, 0), (-1, -1), 0.4, colors.HexColor("#d8d8d8")),
                ("BACKGROUND", (0, 0), (0, -1), colors.HexColor("#f4f5f7")),
                ("VALIGN", (0, 0), (-1, -1), "TOP"),
                ("LEFTPADDING", (0, 0), (-1, -1), 6),
                ("RIGHTPADDING", (0, 0), (-1, -1), 6),
                ("TOPPADDING", (0, 0), (-1, -1), 4),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
            ]
        )
    )
    return table


def _action_sentence(action: dict[str, Any]) -> str:
    message = action.get("message") or action.get("code", "")
    count = action.get("count")
    if isinstance(count, int) and count > 1 and str(count) not in message:
        return f"{message} ({count} ocorrência(s).)"
    return message


def build_pdf_report(report: dict[str, Any], output_path: Path, *, original_filename: str | None = None) -> Path:
    """Constrói o PDF a partir do mesmo dicionário de relatório usado no JSON.

    `report` é o retorno de `abnt_engine.apply_formatting()`: contém
    `config`, `actions_applied`, `issues_remaining` e `manual_validation`.
    """
    styles = _styles()
    output_path.parent.mkdir(parents=True, exist_ok=True)
    doc = SimpleDocTemplate(
        str(output_path),
        pagesize=A4,
        leftMargin=2.5 * cm,
        rightMargin=2.5 * cm,
        topMargin=2 * cm,
        bottomMargin=2 * cm,
        title="Relatório de Revisão ABNT",
    )

    story: list[Any] = []
    story.append(Paragraph("Relatório de Revisão ABNT", styles["title"]))
    filename = original_filename or Path(report.get("input", "documento")).name
    generated_at = datetime.now().strftime("%d/%m/%Y às %H:%M")
    story.append(Paragraph(f"Arquivo revisado: <b>{filename}</b> · Gerado em {generated_at}", styles["subtitle"]))
    story.append(
        Paragraph(
            "Este relatório resume o que o Revisor ABNT alterou automaticamente na sua cópia e o que ainda "
            "precisa de conferência humana. O arquivo original não foi modificado; toda alteração foi aplicada "
            "somente na nova cópia entregue.",
            styles["body"],
        )
    )

    story.append(Paragraph("Configuração usada nesta revisão", styles["h2"]))
    story.append(_config_table(report.get("config", {}), styles))

    actions = report.get("actions_applied") or []
    story.append(Paragraph("O que foi corrigido automaticamente", styles["h2"]))
    if actions:
        items = [ListItem(Paragraph(_action_sentence(action), styles["item"]), bulletColor=colors.HexColor("#1a7a4c")) for action in actions]
        story.append(ListFlowable(items, bulletType="bullet", start="•", leftIndent=14, spaceBefore=2, bulletFontSize=8))
    else:
        story.append(Paragraph("Nenhuma alteração automática foi necessária ou possível neste documento.", styles["empty"]))

    issues = report.get("issues_remaining") or []
    story.append(Paragraph("Pontos que exigem validação humana", styles["h2"]))
    if issues:
        ordered = sorted(issues, key=lambda issue: SEVERITY_ORDER.get(issue.get("severity"), 9))
        items = []
        for issue in ordered:
            label = SEVERITY_LABELS.get(issue.get("severity"), "Informativo")
            text = f"<b>[{label}]</b> {issue.get('message', '')}"
            bullet_color = colors.HexColor("#b3261e") if issue.get("severity") == "error" else (
                colors.HexColor("#8a5300") if issue.get("severity") == "warning" else colors.HexColor("#555555")
            )
            items.append(ListItem(Paragraph(text, styles["item"]), bulletColor=bullet_color))
        story.append(ListFlowable(items, bulletType="bullet", start="•", leftIndent=14, spaceBefore=2, bulletFontSize=8))
    else:
        story.append(Paragraph("Nenhuma pendência identificada pelo motor além das verificações manuais padrão abaixo.", styles["empty"]))

    manual_notes = report.get("manual_validation") or []
    if manual_notes:
        story.append(Paragraph("Sempre confira manualmente antes da entrega", styles["h2"]))
        items = [ListItem(Paragraph(note, styles["item"])) for note in manual_notes]
        story.append(ListFlowable(items, bulletType="bullet", start="•", leftIndent=14, spaceBefore=2, bulletFontSize=8))

    story.append(Spacer(1, 18))
    story.append(
        Paragraph(
            "O Revisor ABNT é uma ferramenta de apoio: não certifica conformidade ABNT nem substitui a conferência "
            "editorial, bibliotecária ou da orientação do curso. O manual da instituição do estudante sempre "
            "prevalece sobre o perfil geral aplicado aqui.",
            styles["footer"],
        )
    )

    doc.build(story)
    return output_path
