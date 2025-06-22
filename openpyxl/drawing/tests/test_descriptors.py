# Copyright (c) 2010-2025 openpyxl


def test_color_descriptor():
    from openpyxl.drawing.colors import ColorChoiceDescriptor

    class DummyStyle:
        value = ColorChoiceDescriptor("value")

    style = DummyStyle()
    style.value = "efefef"
    assert style.value.RGB == "efefef"
