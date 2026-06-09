import unittest

import mover_textos_python as script


class FakeUtility:
    def __init__(self):
        self.prompt_messages = []

    def Prompt(self, message):
        self.prompt_messages.append(message)


class FakeDoc:
    def __init__(self):
        self.Utility = FakeUtility()


class PromptSelectionInAutocadTests(unittest.TestCase):
    def test_prompts_in_autocad_to_select_texts_and_press_enter_in_cad(self):
        doc = FakeDoc()

        script.prompt_selection_in_autocad(doc)

        self.assertEqual(
            doc.Utility.prompt_messages,
            [
                "\nSelecione os textos e polylines no AutoCAD e pressione ENTER no AutoCAD para continuar..."
            ],
        )

    def test_prompts_analysis_status_in_autocad(self):
        doc = FakeDoc()

        script.prompt_analysis_in_autocad(doc)

        self.assertEqual(doc.Utility.prompt_messages, ["\nAnalisando textos selecionados..."])

    def test_prompts_move_instructions_in_autocad(self):
        doc = FakeDoc()

        script.prompt_move_instructions_in_autocad(doc)

        self.assertEqual(
            doc.Utility.prompt_messages,
            [
                "\nMOVIMENTO: W=cima | S=baixo | A=esquerda | D=direita | ESC=proximo conflito"
            ],
        )

    def test_prompts_conflict_choice_commands_in_autocad(self):
        doc = FakeDoc()

        script.prompt_conflict_in_autocad(doc, 1, 3)

        self.assertEqual(
            doc.Utility.prompt_messages,
            [
                "\nCONFLITO 1/3: polylines com area maior que 0 estao se chocando",
                "\nESCOLHA: 1=mover texto 1 | 2=mover texto 2 | P=pular | ESC=finalizar",
            ],
        )


class SelectionSummaryTests(unittest.TestCase):
    def test_builds_summary_with_object_types_and_skip_counts(self):
        lines = script.build_selection_summary(
            selected_count=4,
            accepted_count=1,
            type_counts={"AcDbBlockReference": 2, "AcDbMText": 1, "<erro ao ler tipo>": 1},
            skipped_not_text=2,
            skipped_no_bbox=0,
            skipped_errors=1,
        )

        self.assertIn("Objetos selecionados : 4", lines)
        self.assertIn("  AcDbBlockReference: 2", lines)
        self.assertIn("  AcDbMText: 1", lines)
        self.assertIn("TEXT/MTEXT aceitos   : 1", lines)
        self.assertIn("Ignorados por tipo   : 2", lines)
        self.assertIn("Ignorados por erro   : 1", lines)


class FakeTextObj:
    def __init__(self, color=256):
        self.Color = color
        self.highlight_calls = []
        self.update_calls = 0

    def Highlight(self, enabled):
        self.highlight_calls.append(enabled)

    def Update(self):
        self.update_calls += 1


class FakeDeletableObj:
    def __init__(self):
        self.delete_calls = 0

    def Delete(self):
        self.delete_calls += 1


class ConflictHighlightTests(unittest.TestCase):
    def test_highlight_pair_marks_both_texts_in_cyan(self):
        first = {"obj": FakeTextObj(color=1)}
        second = {"obj": FakeTextObj(color=2)}

        markers = script.highlight_pair(first, second)

        self.assertEqual(markers, [])
        self.assertEqual(first["original_color"], 1)
        self.assertEqual(second["original_color"], 2)
        self.assertEqual(first["obj"].highlight_calls, [True])
        self.assertEqual(second["obj"].highlight_calls, [True])
        self.assertEqual(first["obj"].Color, script.CYAN_COLOR)
        self.assertEqual(second["obj"].Color, script.CYAN_COLOR)


class FakeMarker:
    def __init__(self):
        self.Color = None
        self.Closed = None
        self.update_calls = 0
        self.delete_calls = 0

    def Update(self):
        self.update_calls += 1

    def Delete(self):
        self.delete_calls += 1


class FakeModelSpace:
    def __init__(self):
        self.polylines = []
        self.leaders = []

    def AddLightWeightPolyline(self, coords):
        marker = FakeMarker()
        self.polylines.append((list(coords), marker))
        return marker

    def AddLeader(self, coords, annotation, leader_type):
        marker = FakeMarker()
        self.leaders.append((list(coords), annotation, leader_type, marker))
        return marker


class FakeDocWithModelSpace:
    def __init__(self):
        self.ModelSpace = FakeModelSpace()
        self.regen_calls = []
        self.command_calls = []
        self.cmdnames_calls = 0

    def Regen(self, mode):
        self.regen_calls.append(mode)

    def SendCommand(self, command):
        self.command_calls.append(command)

    def GetVariable(self, name):
        if name == "CMDNAMES":
            self.cmdnames_calls += 1
            return ""

        raise ValueError(name)


class ConflictMarkerTests(unittest.TestCase):
    def test_highlight_pair_only_marks_texts_without_geometry(self):
        doc = FakeDocWithModelSpace()
        first = {"obj": FakeTextObj(color=1), "bbox": (0.0, 0.0, 1.0, 1.0)}
        second = {"obj": FakeTextObj(color=2), "bbox": (10.0, 20.0, 14.0, 22.0)}

        markers = script.highlight_pair(first, second, doc=doc)

        self.assertEqual(markers, [])
        self.assertEqual(doc.ModelSpace.polylines, [])
        self.assertEqual(doc.ModelSpace.leaders, [])
        self.assertEqual(doc.regen_calls, [])

        script.unhighlight_pair(first, second, markers, doc=doc)

        self.assertEqual(first["obj"].Color, 1)
        self.assertEqual(second["obj"].Color, 2)

    def test_tcircle_rectangles_use_offset_constant_width_and_height(self):
        first = {"bbox": (0.0, 0.0, 2.0, 1.0)}
        second = {"bbox": (10.0, 20.0, 14.0, 22.0)}

        first_coords, second_coords = script.tcircle_rectangle_coords_for_pair(first, second)

        self.assertEqual(first_coords, [-1.7, -1.2, 3.7, -1.2, 3.7, 2.2, -1.7, 2.2, -1.7, -1.2])
        self.assertEqual(second_coords, [9.3, 19.3, 14.7, 19.3, 14.7, 22.7, 9.3, 22.7, 9.3, 19.3])

    def test_apply_tcircle_to_pair_draws_rectangles_without_sendcommand(self):
        doc = FakeDocWithModelSpace()
        first = {"bbox": (0.0, 0.0, 2.0, 1.0)}
        second = {"bbox": (10.0, 20.0, 14.0, 22.0)}

        markers = script.apply_tcircle_to_pair(doc, first, second)

        self.assertEqual(doc.command_calls, [])
        self.assertEqual(len(markers), 2)
        self.assertEqual(len(doc.ModelSpace.polylines), 2)
        self.assertEqual(doc.ModelSpace.polylines[0][0], [-1.7, -1.2, 3.7, -1.2, 3.7, 2.2, -1.7, 2.2, -1.7, -1.2])
        self.assertEqual(doc.ModelSpace.polylines[1][0], [9.3, 19.3, 14.7, 19.3, 14.7, 22.7, 9.3, 22.7, 9.3, 19.3])
        self.assertTrue(all(marker.ConstantWidth == 0.0 for marker in markers))

        script.delete_markers(markers, doc=doc)

        self.assertEqual([marker.delete_calls for marker in markers], [1, 1])

    def test_find_conflicts_from_area_polylines_pairs_texts_inside_colliding_polylines(self):
        rect_1 = {
            "handle": "R1",
            "bbox": (0.0, 0.0, 1.0, 1.0),
            "obj": FakeDeletableObj(),
        }
        rect_2 = {
            "handle": "R2",
            "bbox": (0.8, 0.0, 2.0, 1.0),
            "obj": FakeDeletableObj(),
        }
        rect_3 = {
            "handle": "R3",
            "bbox": (10.0, 10.0, 12.0, 12.0),
            "obj": FakeDeletableObj(),
        }
        text_1 = {
            "handle": "T1",
            "bbox": (0.2, 0.2, 0.5, 0.4),
            "object_name": "AcDbMText",
        }
        text_2 = {
            "handle": "T2",
            "bbox": (1.4, 0.2, 1.7, 0.4),
            "object_name": "AcDbText",
        }
        text_3 = {
            "handle": "T3",
            "bbox": (10.2, 10.2, 10.5, 10.4),
            "object_name": "AcDbText",
        }

        conflicts = script.find_conflicts_from_polylines(
            [text_1, text_2, text_3],
            [rect_1, rect_2, rect_3],
        )

        self.assertEqual(len(conflicts), 1)
        self.assertEqual(conflicts[0][0]["handle"], "T1")
        self.assertEqual(conflicts[0][1]["handle"], "T2")
        self.assertIs(conflicts[0][0]["rectangle_obj"], rect_1["obj"])
        self.assertIs(conflicts[0][1]["rectangle_obj"], rect_2["obj"])

    def test_delete_source_rectangle_deletes_rectangle_attached_to_moved_text(self):
        rectangle = FakeDeletableObj()
        item = {"rectangle_obj": rectangle}

        script.delete_source_rectangle(item)

        self.assertEqual(rectangle.delete_calls, 1)

    def test_is_area_polyline_accepts_selected_polylines_with_area_above_zero(self):
        self.assertTrue(script.is_area_polyline("AcDbPolyline", 0.1))
        self.assertTrue(script.is_area_polyline("AcDb2dPolyline", 1.0))
        self.assertFalse(script.is_area_polyline("AcDbPolyline", 0.0))
        self.assertFalse(script.is_area_polyline("AcDbLine", 1.0))

    def test_geometry_conflict_requires_text_rectangles_to_cross(self):
        self.assertTrue(script.geometry_conflict(
            (0.0, 0.0, 2.0, 2.0),
            (1.0, 1.0, 3.0, 3.0),
        ))
        self.assertFalse(script.geometry_conflict(
            (0.0, 0.0, 2.0, 2.0),
            (2.0, 0.0, 4.0, 2.0),
        ))
        self.assertFalse(script.geometry_conflict(
            (0.0, 0.0, 2.0, 2.0),
            (1.0, 2.0, 3.0, 4.0),
        ))
        self.assertFalse(script.geometry_conflict(
            (0.0, 0.0, 2.0, 2.0),
            (2.01, 0.0, 4.0, 2.0),
        ))

    def test_item_geometry_conflict_uses_whole_text_rectangles(self):
        left_text = {
            "bbox": (0.0, 0.0, 4.0, 1.0),
            "geometry_boxes": [(0.0, 0.0, 1.0, 1.0), (3.0, 0.0, 4.0, 1.0)],
            "object_name": "AcDbText",
        }
        inside_empty_space_but_inside_rectangle = {
            "bbox": (1.5, 0.0, 2.5, 1.0),
            "geometry_boxes": [(1.5, 0.0, 2.5, 1.0)],
            "object_name": "AcDbText",
        }
        outside_rectangle = {
            "bbox": (4.1, 0.0, 5.0, 1.0),
            "geometry_boxes": [(4.1, 0.0, 5.0, 1.0)],
            "object_name": "AcDbText",
        }

        self.assertTrue(script.item_geometry_conflict(
            left_text,
            inside_empty_space_but_inside_rectangle,
        ))
        self.assertFalse(script.item_geometry_conflict(left_text, outside_rectangle))


class FakeMTextWithoutNativeBoundingBox:
    ObjectName = "AcDbMText"
    InsertionPoint = (10.0, 20.0, 0.0)
    Width = 5.0
    Height = 0.5
    TextString = "Linha 1\\PLinha 2"
    AttachmentPoint = 1
    Rotation = 0.0

    def GetBoundingBox(self):
        raise RuntimeError("GetBoundingBox unavailable")


class FakeTextWithHeightAndInsertion:
    ObjectName = "AcDbText"
    InsertionPoint = (10.0, 20.0, 0.0)
    Height = 0.14
    TextString = "277"
    Rotation = 0.0

    def GetBoundingBox(self):
        return ((0.0, 0.0, 0.0), (99.0, 99.0, 0.0))


class BoundingBoxTests(unittest.TestCase):
    def test_get_bbox_uses_text_height_and_01_per_character(self):
        bbox = script.get_bbox(FakeTextWithHeightAndInsertion())

        self.assertEqual(bbox, (10.0, 20.0, 10.3, 20.14))

    def test_get_bbox_falls_back_to_mtext_properties_when_native_bbox_fails(self):
        bbox = script.get_bbox(FakeMTextWithoutNativeBoundingBox())

        self.assertEqual(bbox, (10.0, 20.0, 10.7, 20.5))


if __name__ == "__main__":
    unittest.main()
