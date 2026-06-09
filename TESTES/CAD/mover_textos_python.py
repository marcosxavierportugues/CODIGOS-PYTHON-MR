# mover_textos_python.py

from pyautocad import Autocad, APoint, aDouble
from comtypes import COMError
import pythoncom
import keyboard
import math
import re
import time

STEP = 0.1
RECTANGLE_EPSILON = 1e-9
RECTANGLE_PROXIMITY_TOL = 0.04

TEXT_TYPES = {"AcDbText", "AcDbMText"}
POLYLINE_TYPES = {"AcDbPolyline", "AcDb2dPolyline", "AcDb3dPolyline"}

CYAN_COLOR = 4
PINK_COLOR = 6
BYLAYER_COLOR = 256
MARKER_LINE_WIDTH = 0.0
TCIRCLE_OFFSET = 0.35
CHAR_WIDTH = 0.1

RPC_E_CALL_REJECTED = -2147418111
RPC_E_SERVERCALL_RETRYLATER = -2147417846
COM_RETRY_SECONDS = 10
COM_RETRY_DELAY = 0.2


def is_retryable_com_error(error):
    return getattr(error, "hresult", None) in {
        RPC_E_CALL_REJECTED,
        RPC_E_SERVERCALL_RETRYLATER,
    }


def com_retry(action, description="chamada COM"):
    deadline = time.time() + COM_RETRY_SECONDS
    last_error = None

    while time.time() < deadline:
        try:
            return action()
        except COMError as error:
            if not is_retryable_com_error(error):
                raise

            last_error = error
            pythoncom.PumpWaitingMessages()
            time.sleep(COM_RETRY_DELAY)

    raise RuntimeError(
        f"AutoCAD recusou {description} por {COM_RETRY_SECONDS}s. "
        "Feche comandos/janelas abertas no AutoCAD e tente novamente."
    ) from last_error


def wait_for_autocad_idle(doc):
    def is_idle():
        try:
            return str(doc.GetVariable("CMDNAMES")).strip() == ""
        except COMError as error:
            if is_retryable_com_error(error):
                return False
            raise

    deadline = time.time() + COM_RETRY_SECONDS

    while time.time() < deadline:
        if is_idle():
            return

        pythoncom.PumpWaitingMessages()
        time.sleep(COM_RETRY_DELAY)

    raise RuntimeError(
        "O AutoCAD ainda esta com um comando ativo. "
        "Finalize o comando atual e rode o script novamente."
    )


def prompt_in_autocad(doc, message):
    utility = com_retry(lambda: doc.Utility, "acessar prompt do AutoCAD")
    com_retry(
        lambda: utility.Prompt(f"\n{message}"),
        "mostrar mensagem no AutoCAD",
    )


def prompt_selection_in_autocad(doc):
    prompt_in_autocad(
        doc,
        "Selecione os textos e polylines no AutoCAD e pressione ENTER no AutoCAD para continuar...",
    )


def prompt_analysis_in_autocad(doc):
    prompt_in_autocad(doc, "Analisando textos selecionados...")


def prompt_conflict_in_autocad(doc, idx, total):
    prompt_in_autocad(
        doc,
        f"CONFLITO {idx}/{total}: polylines com area maior que 0 estao se chocando",
    )
    prompt_in_autocad(
        doc,
        "ESCOLHA: 1=mover texto 1 | 2=mover texto 2 | P=pular | ESC=finalizar",
    )


def prompt_move_instructions_in_autocad(doc):
    prompt_in_autocad(
        doc,
        "MOVIMENTO: W=cima | S=baixo | A=esquerda | D=direita | ESC=proximo conflito",
    )


def build_selection_summary(
    selected_count,
    accepted_count,
    type_counts,
    skipped_not_text,
    skipped_no_bbox,
    skipped_errors,
):
    lines = [
        "\nResumo da selecao:",
        f"Objetos selecionados : {selected_count}",
    ]

    if type_counts:
        lines.append("Tipos selecionados:")

        for object_name, count in sorted(type_counts.items()):
            lines.append(f"  {object_name}: {count}")

    lines.extend([
        f"TEXT/MTEXT aceitos   : {accepted_count}",
        f"Ignorados por tipo   : {skipped_not_text}",
        f"Ignorados sem bbox   : {skipped_no_bbox}",
        f"Ignorados por erro   : {skipped_errors}",
    ])

    return lines


def get_text(obj):
    try:
        return obj.TextString
    except Exception:
        return ""


def safe_get_attr(obj, attr_name, default=None):
    try:
        return com_retry(lambda: getattr(obj, attr_name), f"ler {attr_name}")
    except Exception:
        return default


def point_xy(point):
    return float(point[0]), float(point[1])


def normalize_mtext_lines(text):
    normalized = str(text or "")
    normalized = normalized.replace("\\P", "\n").replace("\\p", "\n")
    normalized = normalized.replace("\r\n", "\n").replace("\r", "\n")
    return normalized.split("\n") or [""]


def visible_text_lines(text):
    normalized = str(text or "")
    normalized = normalized.replace("\\P", "\n").replace("\\p", "\n")
    normalized = normalized.replace("\r\n", "\n").replace("\r", "\n")
    normalized = re.sub(r"\\[A-Za-z][^;]*;", "", normalized)
    normalized = normalized.replace("{", "").replace("}", "")
    return normalized.split("\n") or [""]


def text_character_count(text):
    lines = visible_text_lines(text)
    if not lines:
        return 0

    return max(len(line) for line in lines)


def rotated_bbox_from_origin(x0, y0, width, height, rotation):
    local_corners = [
        (0.0, 0.0),
        (width, 0.0),
        (width, height),
        (0.0, height),
    ]
    cos_r = math.cos(rotation)
    sin_r = math.sin(rotation)
    corners = []

    for x, y in local_corners:
        corners.append((
            x0 + (x * cos_r - y * sin_r),
            y0 + (x * sin_r + y * cos_r),
        ))

    xs = [point[0] for point in corners]
    ys = [point[1] for point in corners]

    return (
        round(min(xs), 6),
        round(min(ys), 6),
        round(max(xs), 6),
        round(max(ys), 6),
    )


def text_bbox_from_properties(obj):
    insertion = safe_get_attr(obj, "InsertionPoint")
    text_height = float(safe_get_attr(obj, "Height", 0) or 0)

    if not insertion or text_height <= 0:
        return None

    x0, y0 = point_xy(insertion)
    text = safe_get_attr(obj, "TextString", "")
    char_count = text_character_count(text)

    if char_count <= 0:
        return None

    width = char_count * CHAR_WIDTH
    rotation = float(safe_get_attr(obj, "Rotation", 0.0) or 0.0)
    return rotated_bbox_from_origin(x0, y0, width, text_height, rotation)


# bbox formato: (min_x, min_y, min_z, max_x, max_y, max_z)

def text_geometry_boxes_from_bbox(text, bbox):
    lines = visible_text_lines(text)
    min_x, min_y, min_z, max_x, max_y, max_z = bbox
    width = max_x - min_x
    height = max_y - min_y

    if width <= 0 or height <= 0:
        return []

    line_height = height / max(len(lines), 1)
    boxes = []

    for line_index, line in enumerate(lines):
        if not line:
            continue

        char_width = width / max(len(line), 1)
        y_max = max_y - (line_index * line_height)
        y_min = y_max - line_height

        for char_index, char in enumerate(line):
            if char.isspace():
                continue

            x_min = min_x + (char_index * char_width)
            x_max = x_min + char_width
            boxes.append((
                round(x_min, 6),
                round(y_min, 6),
                round(min_z, 6),
                round(x_max, 6),
                round(y_max, 6),
                round(max_z, 6),
            ))

    return boxes


def estimate_mtext_line_count(text, width, text_height):
    lines = normalize_mtext_lines(text)
    line_count = 0
    char_width = max(text_height * 0.6, 0.001)

    for line in lines:
        estimated_width = max(len(line), 1) * char_width
        line_count += max(1, math.ceil(estimated_width / max(width, char_width)))

    return max(line_count, 1)


def mtext_bbox_from_properties(obj):
    return text_bbox_from_properties(obj)


def get_native_bbox(obj):
    try:
        try:
            com_retry(lambda: obj.Update(), "atualizar objeto antes do bounding box")
        except Exception:
            pass

        min_pt, max_pt = com_retry(lambda: obj.GetBoundingBox(), "calcular bounding box")
        return (
            float(min_pt[0]),
            float(min_pt[1]),
            float(max_pt[0]),
            float(max_pt[1]),
        )
    except Exception:
        return None


def get_bbox(obj):
    object_name = safe_get_attr(obj, "ObjectName")

    if object_name in TEXT_TYPES:
        bbox = text_bbox_from_properties(obj)

        if bbox:
            return bbox

    bbox = get_native_bbox(obj)

    if bbox:
        return bbox

    return None


def get_text_geometry_boxes(obj, bbox):
    boxes = text_geometry_boxes_from_bbox(get_text(obj), bbox)
    return boxes or [bbox]


def axis_overlap(min_a, max_a, min_b, max_b):
    return min(max_a, max_b) - max(min_a, min_b)


def geometry_conflict(b1, b2):
    overlap_x = axis_overlap(b1[0], b1[2], b2[0], b2[2])
    overlap_y = axis_overlap(b1[1], b1[3], b2[1], b2[3])
    return overlap_x > RECTANGLE_EPSILON and overlap_y > RECTANGLE_EPSILON


def axis_gap(min_a, max_a, min_b, max_b):
    if max_a < min_b:
        return min_b - max_a

    if max_b < min_a:
        return min_a - max_b

    return 0.0


def rectangle_distance(b1, b2):
    gap_x = axis_gap(b1[0], b1[2], b2[0], b2[2])
    gap_y = axis_gap(b1[1], b1[3], b2[1], b2[3])
    return math.hypot(gap_x, gap_y)


def rectangle_proximity_conflict(b1, b2):
    return rectangle_distance(b1, b2) <= RECTANGLE_PROXIMITY_TOL + RECTANGLE_EPSILON


def bbox_conflict(b1, b2):
    return geometry_conflict(b1, b2)


def bbox_contains(outer, inner):
    return (
        outer[0] <= inner[0] + RECTANGLE_EPSILON and
        outer[1] <= inner[1] + RECTANGLE_EPSILON and
        outer[2] >= inner[2] - RECTANGLE_EPSILON and
        outer[3] >= inner[3] - RECTANGLE_EPSILON
    )


def point_inside_bbox(point, bbox):
    x, y = point
    return (
        bbox[0] - RECTANGLE_EPSILON <= x <= bbox[2] + RECTANGLE_EPSILON and
        bbox[1] - RECTANGLE_EPSILON <= y <= bbox[3] + RECTANGLE_EPSILON
    )


def bbox_center(b):
    return (
        (b[0] + b[2]) / 2.0,
        (b[1] + b[3]) / 2.0,
    )


def order_texts(item_a, item_b):
    ax, ay = bbox_center(item_a["bbox"])
    bx, by = bbox_center(item_b["bbox"])

    if ay > by:
        return item_a, item_b

    if ay < by:
        return item_b, item_a

    if ax <= bx:
        return item_a, item_b

    return item_b, item_a


def move_obj(obj, dx, dy):
    com_retry(
        lambda: obj.Move(APoint(0, 0, 0), APoint(dx, dy, 0)),
        "mover objeto",
    )


def save_original_color(item):
    try:
        item["original_color"] = item["obj"].Color
    except Exception:
        item["original_color"] = BYLAYER_COLOR


def set_color(item, color):
    try:
        item["obj"].Color = color
        item["obj"].Update()
    except Exception:
        pass


def restore_color(item):
    try:
        item["obj"].Color = item.get("original_color", BYLAYER_COLOR)
        item["obj"].Update()
    except Exception:
        pass


def refresh_autocad_view(doc):
    if not doc:
        return

    try:
        com_retry(lambda: doc.Regen(1), "atualizar visualizacao do AutoCAD")
    except Exception:
        pass


def delete_markers(markers, doc=None):
    for marker in markers or []:
        try:
            marker.Delete()
        except Exception:
            pass

    refresh_autocad_view(doc)


def bbox_size(bbox):
    return abs(bbox[2] - bbox[0]), abs(bbox[3] - bbox[1])


def expanded_tcircle_size(bbox, offset_factor=TCIRCLE_OFFSET):
    width, height = bbox_size(bbox)
    offset = max(height, 0.001) * offset_factor

    return width + (2.0 * offset), height + (2.0 * offset)


def fixed_rectangle_coords_around_bbox(bbox, width, height):
    center_x, center_y = bbox_center(bbox)
    half_width = width / 2.0
    half_height = height / 2.0
    min_x = center_x - half_width
    min_y = center_y - half_height
    max_x = center_x + half_width
    max_y = center_y + half_height

    return [
        round(min_x, 6),
        round(min_y, 6),
        round(max_x, 6),
        round(min_y, 6),
        round(max_x, 6),
        round(max_y, 6),
        round(min_x, 6),
        round(max_y, 6),
        round(min_x, 6),
        round(min_y, 6),
    ]


def tcircle_rectangle_coords_for_pair(t1, t2):
    size_1 = expanded_tcircle_size(t1["bbox"])
    size_2 = expanded_tcircle_size(t2["bbox"])
    constant_width = max(size_1[0], size_2[0])
    constant_height = max(size_1[1], size_2[1])

    return (
        fixed_rectangle_coords_around_bbox(t1["bbox"], constant_width, constant_height),
        fixed_rectangle_coords_around_bbox(t2["bbox"], constant_width, constant_height),
    )


def create_polyline_marker(doc, coords, color=PINK_COLOR, line_width=MARKER_LINE_WIDTH):
    model_space = com_retry(lambda: doc.ModelSpace, "acessar ModelSpace")
    marker = com_retry(
        lambda: model_space.AddLightWeightPolyline(aDouble(*coords)),
        "desenhar retangulo",
    )

    try:
        marker.Color = color
        marker.Closed = True
        marker.ConstantWidth = line_width
        marker.Update()
    except Exception:
        pass

    return marker


def apply_tcircle_to_pair(doc, t1, t2):
    coords_1, coords_2 = tcircle_rectangle_coords_for_pair(t1, t2)
    markers = [
        create_polyline_marker(doc, coords_1),
        create_polyline_marker(doc, coords_2),
    ]
    refresh_autocad_view(doc)
    return markers


def is_text_item(item):
    return item.get("object_name") in TEXT_TYPES


def is_area_polyline(object_name, area):
    return object_name in POLYLINE_TYPES and float(area or 0) > 0.0


def item_geometry_conflict(item_a, item_b):
    if not is_text_item(item_a) or not is_text_item(item_b):
        return False

    return geometry_conflict(item_a["bbox"], item_b["bbox"])


def text_inside_rectangle(text_item, rectangle_item):
    text_bbox = text_item["bbox"]
    rectangle_bbox = rectangle_item["bbox"]
    return (
        bbox_contains(rectangle_bbox, text_bbox) or
        point_inside_bbox(bbox_center(text_bbox), rectangle_bbox)
    )


def find_text_for_rectangle(rectangle_item, text_items):
    candidates = [
        item for item in text_items
        if is_text_item(item) and text_inside_rectangle(item, rectangle_item)
    ]

    if not candidates:
        return None

    rect_center = bbox_center(rectangle_item["bbox"])

    def distance_to_rect_center(item):
        text_center = bbox_center(item["bbox"])
        return math.hypot(text_center[0] - rect_center[0], text_center[1] - rect_center[1])

    return min(candidates, key=distance_to_rect_center)


def attach_rectangle_to_text(text_item, rectangle_item):
    item = dict(text_item)
    item["rectangle_obj"] = rectangle_item["obj"]
    item["rectangle_handle"] = rectangle_item["handle"]
    item["rectangle_bbox"] = rectangle_item["bbox"]
    return item


def find_conflicts_from_polylines(text_items, rectangle_items):
    conflicts = []
    used_pairs = set()

    for i in range(len(rectangle_items)):
        for j in range(i + 1, len(rectangle_items)):
            rect_a = rectangle_items[i]
            rect_b = rectangle_items[j]

            if not geometry_conflict(rect_a["bbox"], rect_b["bbox"]):
                continue

            text_a = find_text_for_rectangle(rect_a, text_items)
            text_b = find_text_for_rectangle(rect_b, text_items)

            if not text_a or not text_b or text_a["handle"] == text_b["handle"]:
                continue

            key = tuple(sorted([text_a["handle"], text_b["handle"]]))

            if key in used_pairs:
                continue

            t1, t2 = order_texts(
                attach_rectangle_to_text(text_a, rect_a),
                attach_rectangle_to_text(text_b, rect_b),
            )
            conflicts.append((t1, t2))
            used_pairs.add(key)

    return conflicts


def find_conflicts_from_rectangles(text_items, rectangle_items):
    return find_conflicts_from_polylines(text_items, rectangle_items)


def delete_source_rectangle(item):
    rectangle = item.get("rectangle_obj")

    if not rectangle:
        return

    try:
        rectangle.Delete()
    except Exception:
        pass


def highlight_pair(t1, t2, doc=None):
    save_original_color(t1)
    save_original_color(t2)

    try:
        t1["obj"].Highlight(True)
        t2["obj"].Highlight(True)
    except Exception:
        pass

    set_color(t1, CYAN_COLOR)
    set_color(t2, CYAN_COLOR)

    return []


def unhighlight_pair(t1, t2, markers=None, doc=None):
    delete_markers(markers, doc=doc)
    restore_color(t1)
    restore_color(t2)

    try:
        t1["obj"].Highlight(False)
        t2["obj"].Highlight(False)
    except Exception:
        pass


def select_area(acad):
    doc = com_retry(lambda: acad.doc, "acessar documento ativo")
    wait_for_autocad_idle(doc)
    name = "PY_TEXT_CONFLICT_SELECTION"

    try:
        old = com_retry(
            lambda: doc.SelectionSets.Item(name),
            "buscar selecao anterior",
        )
        com_retry(lambda: old.Delete(), "apagar selecao anterior")
    except Exception:
        pass

    ss = com_retry(lambda: doc.SelectionSets.Add(name), "criar selecao")

    print("\nSelecione os textos e polylines no AutoCAD.")
    prompt_selection_in_autocad(doc)
    com_retry(lambda: ss.SelectOnScreen(), "selecionar na tela")
    prompt_analysis_in_autocad(doc)
    print("\nAnalisando textos selecionados...")

    text_items = []
    rectangle_items = []
    selected_count = 0
    skipped_not_text = 0
    skipped_no_bbox = 0
    skipped_errors = 0
    type_counts = {}

    for obj in ss:
        selected_count += 1

        try:
            object_name = com_retry(lambda: obj.ObjectName, "ler tipo do objeto")
            type_counts[object_name] = type_counts.get(object_name, 0) + 1
            area = safe_get_attr(obj, "Area", 0)

            if object_name in TEXT_TYPES:
                bbox = get_bbox(obj)

                if not bbox:
                    skipped_no_bbox += 1
                    continue

                text_items.append({
                    "obj": obj,
                    "text": get_text(obj),
                    "bbox": bbox,
                    "handle": obj.Handle,
                    "object_name": object_name,
                    "original_color": BYLAYER_COLOR,
                })
                continue

            if is_area_polyline(object_name, area):
                bbox = get_native_bbox(obj)

                if not bbox:
                    skipped_no_bbox += 1
                    continue

                rectangle_items.append({
                    "obj": obj,
                    "bbox": bbox,
                    "handle": obj.Handle,
                    "object_name": object_name,
                    "area": float(area or 0),
                })
                continue

            if object_name not in TEXT_TYPES:
                skipped_not_text += 1
                continue

        except Exception:
            skipped_errors += 1
            type_counts["<erro ao ler tipo>"] = type_counts.get("<erro ao ler tipo>", 0) + 1
            continue

    for line in build_selection_summary(
        selected_count=selected_count,
        accepted_count=len(text_items),
        type_counts=type_counts,
        skipped_not_text=skipped_not_text,
        skipped_no_bbox=skipped_no_bbox,
        skipped_errors=skipped_errors,
    ):
        print(line)

    print(f"Polylines com area maior que 0: {len(rectangle_items)}")

    com_retry(lambda: ss.Delete(), "apagar selecao temporaria")
    return text_items, rectangle_items


def find_conflicts(items):
    conflicts = []
    used_pairs = set()

    for i in range(len(items)):
        for j in range(i + 1, len(items)):
            a = items[i]
            b = items[j]

            if item_geometry_conflict(a, b):
                key = tuple(sorted([a["handle"], b["handle"]]))

                if key not in used_pairs:
                    t1, t2 = order_texts(a, b)
                    conflicts.append((t1, t2))
                    used_pairs.add(key)

    return conflicts


def move_mode(obj, doc=None):
    print("\nModo movimentação ativo")
    print("W = cima")
    print("S = baixo")
    print("A = esquerda")
    print("D = direita")
    print("ESC = próximo conflito")

    if doc:
        prompt_move_instructions_in_autocad(doc)

    while True:
        key = str(keyboard.read_key(suppress=True)).lower()

        if key == "esc":
            print("\nMovimentação encerrada.")
            time.sleep(0.25)
            return

        if key == "w":
            move_obj(obj, 0, STEP)
        elif key == "s":
            move_obj(obj, 0, -STEP)
        elif key == "a":
            move_obj(obj, -STEP, 0)
        elif key == "d":
            move_obj(obj, STEP, 0)

        time.sleep(0.05)


def main():
    pythoncom.CoInitialize()

    acad = Autocad(create_if_not_exists=False)
    print("\nConectado ao AutoCAD.")
    doc = com_retry(lambda: acad.doc, "acessar documento ativo")

    text_items, rectangle_items = select_area(acad)

    if not text_items:
        print("\nNenhum TEXT ou MTEXT encontrado.")
        return

    if not rectangle_items:
        print("\nNenhuma polyline com area maior que 0 encontrada.")
        return

    prompt_analysis_in_autocad(doc)
    print("\nAnalisando conflitos entre polylines com area maior que 0...")
    conflicts = find_conflicts_from_polylines(text_items, rectangle_items)

    print(f"\nTextos analisados: {len(text_items)}")
    print(f"Polylines analisadas: {len(rectangle_items)}")
    print(f"Conflitos encontrados: {len(conflicts)}")

    if not conflicts:
        print("\nNenhum conflito encontrado.")
        return

    resolved = 0
    skipped = 0
    esc_count = 0
    for idx, (t1, t2) in enumerate(conflicts, start=1):
        markers = highlight_pair(t1, t2, doc=doc)
        prompt_conflict_in_autocad(doc, idx, len(conflicts))

        print("\n===================================")
        print(f"CONFLITO {idx} / {len(conflicts)}")

        print("\nTexto 1")
        print(f"Handle: {t1['handle']}")
        print(f"Texto : {t1['text']}")

        print("\nTexto 2")
        print(f"Handle: {t2['handle']}")
        print(f"Texto : {t2['text']}")

        print("\nOs dois textos do conflito estao selecionados/destacados em CYAN no AutoCAD.")
        print("\nEscolha:")
        print("1 = mover texto 1")
        print("2 = mover texto 2")
        print("P = pular")
        print("ESC = finalizar se apertar duas vezes")

        choice = str(keyboard.read_key(suppress=True)).lower()
        time.sleep(0.25)

        if choice == "esc":
            esc_count += 1

            if esc_count >= 2:
                unhighlight_pair(t1, t2, markers, doc=doc)
                print("\nFinalizado por ESC duplo.")
                break

            print("\nESC pressionado.")
            print("Pressione ESC novamente para finalizar.")
            skipped += 1
            unhighlight_pair(t1, t2, markers, doc=doc)
            continue

        esc_count = 0

        if choice == "1":
            print("\nMovendo TEXTO 1")
            move_mode(t1["obj"], doc=doc)
            delete_source_rectangle(t1)
            resolved += 1

        elif choice == "2":
            print("\nMovendo TEXTO 2")
            move_mode(t2["obj"], doc=doc)
            delete_source_rectangle(t2)
            resolved += 1

        elif choice == "p":
            print("\nConflito pulado.")
            skipped += 1

        else:
            print("\nOpção inválida.")
            skipped += 1

        unhighlight_pair(t1, t2, markers, doc=doc)

    print("\n===================================")
    print("FINALIZADO")
    print(f"Conflitos encontrados : {len(conflicts)}")
    print(f"Conflitos tratados    : {resolved}")
    print(f"Conflitos pulados     : {skipped}")


if __name__ == "__main__":
    main()
