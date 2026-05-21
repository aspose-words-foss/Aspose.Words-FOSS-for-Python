"""Visible-run filter: drop instruction runs inside a field's code section."""


from aspose.words_foss import light_document_model as ldm


def visible_runs(para: ldm.Paragraph) -> list[ldm.Run]:
    """Return *para*'s runs minus the ones inside a field's code section."""
    code_depth = 0
    out: list[ldm.Run] = []
    for child in para._children:
        if isinstance(child, ldm.FieldStart):
            code_depth += 1
        elif isinstance(child, ldm.FieldSeparator):
            if code_depth > 0:
                code_depth -= 1
        elif isinstance(child, ldm.FieldEnd):
            if code_depth > 0:
                code_depth -= 1
        elif isinstance(child, ldm.Run) and code_depth == 0:
            out.append(child)
    return out
