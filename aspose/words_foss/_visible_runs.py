"""Visible-run filter: drop instruction runs inside a field's code section."""


from aspose.words_foss import light_document_model as ldm

# ShapeType.RECTANGLE — Shape.CreateHorizontalRule builds a filled inline
# rectangle, and nothing else in the pipeline assigns this shape type.
HORIZONTAL_RULE_SHAPE_TYPE = 1


def is_horizontal_rule_shape(shape: ldm.Shape) -> bool:
    """Whether *shape* is the rectangle a horizontal rule is drawn with."""
    return (
        shape.shape_type == HORIZONTAL_RULE_SHAPE_TYPE
        and not shape.has_image
        and shape.text_box is None
    )


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
