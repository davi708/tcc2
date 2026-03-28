import hashlib
import io
import json
import math
import secrets
import ssl
import tempfile
import urllib.parse
import urllib.request
from collections import Counter, defaultdict
from pathlib import Path
from typing import Any, Dict, List, Tuple

import ezdxf
from PIL import Image, ImageDraw
import pandas as pd
import plotly.graph_objects as go
import streamlit as st
import streamlit.components.v1 as components
import streamlit.elements.image as st_image
from streamlit.elements.lib.image_utils import image_to_url as _streamlit_image_to_url
from streamlit.elements.lib.layout_utils import LayoutConfig
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter
from streamlit_plotly_events import plotly_events
from streamlit_drawable_canvas import st_canvas

try:
    import certifi
except ModuleNotFoundError:
    certifi = None

try:
    import tomllib
except ModuleNotFoundError:
    tomllib = None

st.set_page_config(page_title="Dimensionamento Eletrico Residencial", layout="wide")

if not hasattr(st_image, "image_to_url"):
    def _compat_image_to_url(image, width, clamp, channels, output_format, image_id):
        return _streamlit_image_to_url(
            image=image,
            layout_config=LayoutConfig(width=width),
            clamp=clamp,
            channels=channels,
            output_format=output_format,
            image_id=image_id,
        )

    st_image.image_to_url = _compat_image_to_url


def _session_key(prefix: str) -> str:
    return f"dxf_import_{prefix}"


def _hash_bytes(data: bytes) -> str:
    return hashlib.sha256(data).hexdigest()


def _extract_segments(doc: ezdxf.EzDxfDocument) -> list[tuple[tuple[float, float], tuple[float, float]]]:
    msp = doc.modelspace()
    segments: list[tuple[tuple[float, float], tuple[float, float]]] = []

    for entity in msp.query("LINE"):
        start = (float(entity.dxf.start.x), float(entity.dxf.start.y))
        end = (float(entity.dxf.end.x), float(entity.dxf.end.y))
        if start != end:
            segments.append((start, end))

    for entity in msp.query("LWPOLYLINE"):
        points = [(float(x), float(y)) for x, y, *_ in entity.get_points("xy")]
        if entity.closed and points and points[0] != points[-1]:
            points.append(points[0])
        for i in range(len(points) - 1):
            if points[i] != points[i + 1]:
                segments.append((points[i], points[i + 1]))

    for entity in msp.query("POLYLINE"):
        if entity.get_mode() != "AcDb2dPolyline":
            continue
        points = [(float(vertex.dxf.location.x), float(vertex.dxf.location.y)) for vertex in entity.vertices]
        if entity.is_closed and points and points[0] != points[-1]:
            points.append(points[0])
        for i in range(len(points) - 1):
            if points[i] != points[i + 1]:
                segments.append((points[i], points[i + 1]))

    return segments


def _load_dxf_payload(file_bytes: bytes) -> tuple[list[tuple[tuple[float, float], tuple[float, float]]], list[tuple[float, float]]]:
    tmp_path: Path | None = None
    try:
        with tempfile.NamedTemporaryFile(delete=False, suffix=".dxf") as tmp:
            tmp.write(file_bytes)
            tmp_path = Path(tmp.name)
        doc = ezdxf.readfile(tmp_path)
        segments = _extract_segments(doc)
        if not segments:
            raise ValueError("Nenhuma entidade grafica utilizavel foi encontrada no DXF.")
        points = sorted({point for segment in segments for point in segment})
        return segments, points
    finally:
        if tmp_path and tmp_path.exists():
            try:
                tmp_path.unlink()
            except OSError:
                pass


def _polygon_area(points: list[tuple[float, float]]) -> float:
    acc = 0.0
    total = len(points)
    for i in range(total):
        x1, y1 = points[i]
        x2, y2 = points[(i + 1) % total]
        acc += x1 * y2 - x2 * y1
    return abs(acc) / 2.0


def _polygon_perimeter(points: list[tuple[float, float]]) -> float:
    total = 0.0
    count = len(points)
    for i in range(count):
        x1, y1 = points[i]
        x2, y2 = points[(i + 1) % count]
        total += math.hypot(x2 - x1, y2 - y1)
    return total


def _polygon_centroid(points: list[tuple[float, float]]) -> tuple[float, float]:
    area_factor = 0.0
    cx = 0.0
    cy = 0.0
    count = len(points)
    for i in range(count):
        x1, y1 = points[i]
        x2, y2 = points[(i + 1) % count]
        cross = x1 * y2 - x2 * y1
        area_factor += cross
        cx += (x1 + x2) * cross
        cy += (y1 + y2) * cross

    if abs(area_factor) < 1e-9:
        xs = [point[0] for point in points]
        ys = [point[1] for point in points]
        return sum(xs) / len(xs), sum(ys) / len(ys)

    area_factor *= 0.5
    cx /= 6.0 * area_factor
    cy /= 6.0 * area_factor
    return cx, cy


def _distance(a: tuple[float, float], b: tuple[float, float]) -> float:
    return math.hypot(a[0] - b[0], a[1] - b[1])


def _nearest_endpoint(clicked: tuple[float, float], endpoints: list[tuple[float, float]]) -> tuple[float, float]:
    return min(endpoints, key=lambda point: _distance(point, clicked))


def _constrain_to_orthogonal(
    clicked: tuple[float, float],
    last_point: tuple[float, float],
) -> tuple[float, float]:
    dx = clicked[0] - last_point[0]
    dy = clicked[1] - last_point[1]
    if abs(dx) >= abs(dy):
        return clicked[0], last_point[1]
    return last_point[0], clicked[1]


DXF_MARGIN_FACTOR = 0.03


def _apply_dxf_margin(valor: float) -> float:
    return math.ceil((valor * (1.0 + DXF_MARGIN_FACTOR)) * 10.0) / 10.0


def _infer_tipo_comodo(nome: str) -> str:
    nome_normalizado = normalizar_ambiente(nome)
    nome_bruto = nome.strip().lower()

    if 'lavabo' in nome_bruto:
        return 'lavabo'
    if 'banheiro' in nome_bruto or 'banho' in nome_bruto or nome_normalizado == 'banheiro':
        return 'banheiro'
    if 'suite' in nome_bruto or nome_normalizado == 'suite':
        return 'suite'
    if 'dorm' in nome_bruto or 'quarto' in nome_bruto or nome_normalizado == 'quarto':
        return 'quarto'
    if 'cozinha' in nome_bruto or nome_normalizado == 'cozinha':
        return 'cozinha'
    if 'serv' in nome_bruto:
        return 'area_servico'
    if 'varanda' in nome_bruto:
        return 'varanda'
    if 'corredor' in nome_bruto:
        return 'corredor'
    if 'circ' in nome_bruto or nome_normalizado == 'circulacao':
        return 'circulacao'
    if 'hall' in nome_bruto:
        return 'hall'
    if 'jantar' in nome_bruto or 'estar' in nome_bruto or nome_normalizado == 'sala':
        return 'sala'
    return nome_normalizado if nome_normalizado else 'outro'


def _build_imported_room(points: list[tuple[float, float]], index: int) -> dict[str, Any]:
    area_bruta = _polygon_area(points)
    perimetro_bruto = _polygon_perimeter(points)
    centroide_x, centroide_y = _polygon_centroid(points)
    area = _apply_dxf_margin(area_bruta)
    perimetro = _apply_dxf_margin(perimetro_bruto)
    return {
        "nome": f"Comodo {index}",
        "tipo_sugerido": 'outro',
        "area": area,
        "perimetro": perimetro,
        "area_bruta": area_bruta,
        "perimetro_bruto": perimetro_bruto,
        "centroide_x": centroide_x,
        "centroide_y": centroide_y,
        "vertices": points[:],
    }


def _project_bounds(
    endpoints: list[tuple[float, float]],
) -> tuple[float, float, float, float, float, float]:
    xs = [point[0] for point in endpoints]
    ys = [point[1] for point in endpoints]
    min_x, max_x = min(xs), max(xs)
    min_y, max_y = min(ys), max(ys)
    pad_x = max((max_x - min_x) * 0.04, 1.0)
    pad_y = max((max_y - min_y) * 0.04, 1.0)
    return min_x, max_x, min_y, max_y, pad_x, pad_y


def _default_view_state(endpoints: list[tuple[float, float]]) -> dict[str, float]:
    min_x, max_x, min_y, max_y, pad_x, pad_y = _project_bounds(endpoints)
    return {
        'center_x': (min_x + max_x) / 2.0,
        'center_y': (min_y + max_y) / 2.0,
        'zoom': 1.0,
        'min_x': min_x,
        'max_x': max_x,
        'min_y': min_y,
        'max_y': max_y,
        'pad_x': pad_x,
        'pad_y': pad_y,
    }


def _clamp_view_state(view_state: dict[str, float]) -> dict[str, float]:
    min_x = view_state['min_x'] - view_state['pad_x']
    max_x = view_state['max_x'] + view_state['pad_x']
    min_y = view_state['min_y'] - view_state['pad_y']
    max_y = view_state['max_y'] + view_state['pad_y']
    full_width = max(max_x - min_x, 1.0)
    full_height = max(max_y - min_y, 1.0)

    zoom = min(max(view_state.get('zoom', 1.0), 1.0), 8.0)
    visible_width = full_width / zoom
    visible_height = full_height / zoom

    half_w = visible_width / 2.0
    half_h = visible_height / 2.0

    center_x = min(max(view_state.get('center_x', (min_x + max_x) / 2.0), min_x + half_w), max_x - half_w)
    center_y = min(max(view_state.get('center_y', (min_y + max_y) / 2.0), min_y + half_h), max_y - half_h)

    return {
        **view_state,
        'center_x': center_x,
        'center_y': center_y,
        'zoom': zoom,
    }


def _update_view_state(view_state: dict[str, float], action: str) -> dict[str, float]:
    state = _clamp_view_state(view_state)
    min_x = state['min_x'] - state['pad_x']
    max_x = state['max_x'] + state['pad_x']
    min_y = state['min_y'] - state['pad_y']
    max_y = state['max_y'] + state['pad_y']
    full_width = max(max_x - min_x, 1.0)
    full_height = max(max_y - min_y, 1.0)
    visible_width = full_width / state['zoom']
    visible_height = full_height / state['zoom']
    move_x = visible_width * 0.12
    move_y = visible_height * 0.12

    updated = dict(state)
    if action == 'zoom_in':
        updated['zoom'] = min(state['zoom'] * 1.25, 8.0)
    elif action == 'zoom_out':
        updated['zoom'] = max(state['zoom'] / 1.25, 1.0)
    elif action == 'left':
        updated['center_x'] = state['center_x'] - move_x
    elif action == 'right':
        updated['center_x'] = state['center_x'] + move_x
    elif action == 'up':
        updated['center_y'] = state['center_y'] + move_y
    elif action == 'down':
        updated['center_y'] = state['center_y'] - move_y
    elif action == 'reset':
        updated['center_x'] = (state['min_x'] + state['max_x']) / 2.0
        updated['center_y'] = (state['min_y'] + state['max_y']) / 2.0
        updated['zoom'] = 1.0
    return _clamp_view_state(updated)


def _build_canvas_background(
    segments: list[tuple[tuple[float, float], tuple[float, float]]],
    endpoints: list[tuple[float, float]],
    view_state: dict[str, float],
    max_width: int = 1500,
    max_height: int = 980,
) -> tuple[Image.Image, dict[str, float]]:
    state = _clamp_view_state(view_state)
    min_x = state['min_x'] - state['pad_x']
    max_x = state['max_x'] + state['pad_x']
    min_y = state['min_y'] - state['pad_y']
    max_y = state['max_y'] + state['pad_y']
    full_width = max(max_x - min_x, 1.0)
    full_height = max(max_y - min_y, 1.0)

    visible_width = full_width / state['zoom']
    visible_height = full_height / state['zoom']
    left = state['center_x'] - visible_width / 2.0
    right = state['center_x'] + visible_width / 2.0
    bottom = state['center_y'] - visible_height / 2.0
    top = state['center_y'] + visible_height / 2.0

    canvas_width = max_width
    canvas_height = max_height
    scale = min(canvas_width / visible_width, canvas_height / visible_height)
    draw_width = visible_width * scale
    draw_height = visible_height * scale
    offset_x = (canvas_width - draw_width) / 2.0
    offset_y = (canvas_height - draw_height) / 2.0

    image = Image.new('RGB', (canvas_width, canvas_height), '#1f2937')
    draw = ImageDraw.Draw(image)

    def to_canvas(point: tuple[float, float]) -> tuple[float, float]:
        x = (point[0] - left) * scale + offset_x
        y = (top - point[1]) * scale + offset_y
        return x, y

    for start_point, end_point in segments:
        x1, y1 = to_canvas(start_point)
        x2, y2 = to_canvas(end_point)
        draw.line((x1, y1, x2, y2), fill='#e5e7eb', width=2)

    transform = {
        'left': left,
        'top': top,
        'scale': scale,
        'offset_x': offset_x,
        'offset_y': offset_y,
        'canvas_width': float(canvas_width),
        'canvas_height': float(canvas_height),
    }
    return image, transform


def _canvas_to_world_point(point: tuple[float, float], transform: dict[str, float]) -> tuple[float, float]:
    x = (point[0] - transform['offset_x']) / transform['scale'] + transform['left']
    y = transform['top'] - ((point[1] - transform['offset_y']) / transform['scale'])
    return round(x, 4), round(y, 4)


def _extract_polygon_points_from_canvas_object(obj: dict[str, Any]) -> list[tuple[float, float]]:
    if obj.get('type') == 'path' and obj.get('path'):
        left = float(obj.get('left', 0.0))
        top = float(obj.get('top', 0.0))
        path_offset = obj.get('pathOffset') or {}
        offset_x = float(path_offset.get('x', 0.0))
        offset_y = float(path_offset.get('y', 0.0))
        points: list[tuple[float, float]] = []
        for command in obj.get('path', []):
            if not command:
                continue
            opcode = command[0]
            if opcode in ('M', 'L') and len(command) >= 3:
                x = float(command[1]) + left - offset_x
                y = float(command[2]) + top - offset_y
                point = (x, y)
                if not points or point != points[-1]:
                    points.append(point)
        if len(points) > 1 and points[0] == points[-1]:
            points.pop()
        return points

    if obj.get('type') == 'polygon' and obj.get('points'):
        left = float(obj.get('left', 0.0))
        top = float(obj.get('top', 0.0))
        path_offset = obj.get('pathOffset') or {}
        offset_x = float(path_offset.get('x', 0.0))
        offset_y = float(path_offset.get('y', 0.0))
        points = []
        for point in obj.get('points', []):
            x = float(point.get('x', 0.0)) + left - offset_x
            y = float(point.get('y', 0.0)) + top - offset_y
            points.append((x, y))
        return points

    return []


def _extract_latest_room_from_canvas(
    canvas_json: dict[str, Any] | None,
    transform: dict[str, float],
    index: int,
) -> dict[str, Any] | None:
    if not canvas_json:
        return None
    objects = canvas_json.get('objects') or []
    for obj in reversed(objects):
        points_canvas = _extract_polygon_points_from_canvas_object(obj)
        if len(points_canvas) >= 3:
            points_world = [_canvas_to_world_point(point, transform) for point in points_canvas]
            return _build_imported_room(points_world, index)
    return None


def _inject_dxf_hotkeys(prefix: str) -> None:
    components.html(
        f"""
        <script>
        (() => {{
          const flag = '__dxf_hotkeys_{prefix}__';
          if (window.parent[flag]) return;
          window.parent[flag] = true;

          const findParentButton = (label) => Array.from(window.parent.document.querySelectorAll('button')).find((btn) => (btn.innerText || '').trim() === label);
          const clickToolbarByAlt = (altText) => {{
            const frames = Array.from(window.parent.document.querySelectorAll('iframe'));
            for (const frame of frames) {{
              try {{
                const doc = frame.contentDocument || frame.contentWindow.document;
                if (!doc) continue;
                const img = doc.querySelector(`img[alt="${{altText}}"]`);
                if (img) {{
                  const button = img.closest('button') || img.parentElement;
                  if (button) {{ button.click(); return true; }}
                }}
              }} catch (err) {{}}
            }}
            return false;
          }};

          const handler = (event) => {{
            const tag = (event.target && event.target.tagName) ? event.target.tagName.toLowerCase() : '';
            if (tag === 'input' || tag === 'textarea') return;

            if (event.key === 'Enter') {{
              const btn = findParentButton('Salvar comodo desenhado');
              if (btn) {{ event.preventDefault(); btn.click(); }}
            }}
            if (event.key === 'Escape') {{
              const btn = findParentButton('Limpar desenho atual');
              if (btn) {{ event.preventDefault(); btn.click(); }}
            }}
            if ((event.ctrlKey || event.metaKey) && event.key.toLowerCase() === 'z') {{
              if (clickToolbarByAlt('Undo last operation')) {{ event.preventDefault(); }}
            }}
            if ((event.ctrlKey || event.metaKey) && event.key.toLowerCase() === 'y') {{
              if (clickToolbarByAlt('Redo last operation')) {{ event.preventDefault(); }}
            }}
            if (event.key === '+') {{
              const btn = findParentButton('Zoom +');
              if (btn) {{ event.preventDefault(); btn.click(); }}
            }}
            if (event.key === '-') {{
              const btn = findParentButton('Zoom -');
              if (btn) {{ event.preventDefault(); btn.click(); }}
            }}
          }};

          window.parent.document.addEventListener('keydown', handler, true);
          window.document.addEventListener('keydown', handler, true);
          const frames = Array.from(window.parent.document.querySelectorAll('iframe'));
          for (const frame of frames) {{
            try {{
              const doc = frame.contentDocument || frame.contentWindow.document;
              if (doc) doc.addEventListener('keydown', handler, true);
            }} catch (err) {{}}
          }}
        }})();
        </script>
        """,
        height=0,
    )


def _render_dxf_editor_panel(
    *,
    segments: list[tuple[tuple[float, float], tuple[float, float]]],
    endpoints: list[tuple[float, float]],
    rooms_key: str,
    canvas_key: str,
    open_key: str,
    view_key: str,
    pending_key: str,
    pending_name_key: str,
    dialog_mode: bool,
) -> None:
    rooms = st.session_state[rooms_key]
    canvas_state = st.session_state.get(canvas_key)
    view_state = st.session_state.get(view_key, _default_view_state(endpoints))
    background_image, transform = _build_canvas_background(segments, endpoints, view_state)

    _inject_dxf_hotkeys('canvas')

    st.markdown(
        """
        <style>
        iframe[title="streamlit_drawable_canvas.st_canvas"] {
            width: 100% !important;
        }
        </style>
        """,
        unsafe_allow_html=True,
    )

    pending_room = st.session_state.get(pending_key)
    if pending_room:
        with st.container(border=True):
            st.markdown('**Nomear c?modo**')
            st.write(f"Area: {pending_room['area']:.2f} m2 | Perimetro: {pending_room['perimetro']:.2f} m")
            nome_digitado = st.text_input('Qual o nome do c?modo?', key=pending_name_key, value=st.session_state.get(pending_name_key, ''))
            confirm_col, cancel_col = st.columns(2)
            with confirm_col:
                confirmar_nome = st.button('Confirmar nome do c?modo', use_container_width=True, type='primary')
            with cancel_col:
                cancelar_nome = st.button('Cancelar salvamento', use_container_width=True)

            if confirmar_nome:
                nome_limpo = nome_digitado.strip()
                if not nome_limpo:
                    st.warning('Digite um nome para o c?modo antes de confirmar.')
                else:
                    pending_room['nome'] = nome_limpo
                    pending_room['tipo_sugerido'] = _infer_tipo_comodo(nome_limpo)
                    rooms.append(pending_room)
                    st.session_state[rooms_key] = rooms
                    st.session_state[pending_key] = None
                    st.session_state[pending_name_key] = ''
                    st.session_state[canvas_key] = None
                    st.success(f"{nome_limpo} salvo com sucesso.")
                    st.rerun()

            if cancelar_nome:
                st.session_state[pending_key] = None
                st.session_state[pending_name_key] = ''
                st.rerun()

    nav_cols = st.columns([1, 1, 1, 1, 1, 1, 1, 1] if dialog_mode else [1, 1, 1, 1, 1, 1, 1])
    nav_prefix = _session_key('dialog_nav' if dialog_mode else 'inline_nav')
    with nav_cols[0]:
        zoom_in = st.button('Zoom +', key=f'{nav_prefix}_zoom_in', use_container_width=True)
    with nav_cols[1]:
        zoom_out = st.button('Zoom -', key=f'{nav_prefix}_zoom_out', use_container_width=True)
    with nav_cols[2]:
        move_left = st.button('Esq', key=f'{nav_prefix}_left', use_container_width=True)
    with nav_cols[3]:
        move_right = st.button('Dir', key=f'{nav_prefix}_right', use_container_width=True)
    with nav_cols[4]:
        move_up = st.button('Cima', key=f'{nav_prefix}_up', use_container_width=True)
    with nav_cols[5]:
        move_down = st.button('Baixo', key=f'{nav_prefix}_down', use_container_width=True)
    with nav_cols[6]:
        reset_view = st.button('Centralizar', key=f'{nav_prefix}_reset', use_container_width=True)
    fechar_editor = False
    if dialog_mode:
        with nav_cols[7]:
            fechar_editor = st.button('Fechar editor ampliado', key=f'{nav_prefix}_close', use_container_width=True)

    action = None
    if zoom_in:
        action = 'zoom_in'
    elif zoom_out:
        action = 'zoom_out'
    elif move_left:
        action = 'left'
    elif move_right:
        action = 'right'
    elif move_up:
        action = 'up'
    elif move_down:
        action = 'down'
    elif reset_view:
        action = 'reset'

    if action:
        st.session_state[view_key] = _update_view_state(view_state, action)
        st.session_state[canvas_key] = None
        st.rerun()

    top_cols = st.columns([1.1, 1.0, 1.1, 1.2] if dialog_mode else [1.1, 1.0, 1.1])
    action_prefix = _session_key('dialog_action' if dialog_mode else 'inline_action')
    with top_cols[0]:
        salvar = st.button('Salvar comodo desenhado', key=f'{action_prefix}_save', use_container_width=True)
    with top_cols[1]:
        limpar = st.button('Limpar desenho atual', key=f'{action_prefix}_clear', use_container_width=True)
    with top_cols[2]:
        redefinir = st.button('Redefinir comodos do DXF', key=f'{action_prefix}_reset_rooms', use_container_width=True)

    if redefinir:
        st.session_state[rooms_key] = []
        st.session_state[canvas_key] = None
        st.session_state[pending_key] = None
        st.session_state[pending_name_key] = ''
        st.rerun()

    if limpar:
        st.session_state[canvas_key] = None
        st.session_state[pending_key] = None
        st.session_state[pending_name_key] = ''
        st.rerun()

    canvas_result = st_canvas(
        fill_color='rgba(16, 185, 129, 0.22)',
        stroke_width=2,
        stroke_color='#22c55e',
        background_image=background_image,
        update_streamlit=True,
        height=int(transform['canvas_height']),
        width=int(transform['canvas_width']),
        drawing_mode='polygon',
        initial_drawing=canvas_state,
        display_toolbar=True,
        point_display_radius=3,
        key=_session_key('canvas_dialog' if dialog_mode else 'canvas_inline'),
    )

    if canvas_result.json_data is not None:
        st.session_state[canvas_key] = canvas_result.json_data
        canvas_state = canvas_result.json_data

    if salvar and not pending_room:
        novo = _extract_latest_room_from_canvas(canvas_state, transform, len(rooms) + 1)
        if novo is None:
            st.warning('Desenhe um poligono fechado no canvas antes de salvar o comodo. Feche com duplo clique ou clique direito.')
        else:
            st.session_state[pending_key] = novo
            st.session_state[pending_name_key] = ''
            st.rerun()

    if fechar_editor:
        st.session_state[open_key] = False
        st.rerun()

    st.caption('Clique nos vertices do c?modo sobre a planta. Feche o pol?gono com duplo clique ou clique direito.')
    st.caption('Atalhos: Enter salva, Esc limpa, Ctrl+Z desfaz, Ctrl+Y refaz, + e - controlam o zoom.')

    info_col1, info_col2 = st.columns([1.5, 1])
    with info_col1:
        st.markdown('**Como usar**')
        st.write('1. Ajuste a vista com Zoom e setas, se precisar.')
        st.write('2. Clique nos cantos do c?modo em ordem.')
        st.write('3. Feche com duplo clique ou bot?o direito.')
        st.write('4. Clique em Salvar c?modo desenhado e informe o nome.')
    with info_col2:
        st.markdown('**Comodos importados**')
        if not rooms:
            st.info('Nenhum comodo salvo ainda.')
        else:
            for idx, room in enumerate(rooms, start=1):
                st.markdown(
                    f"**{idx}. {room['nome']}**  \nArea: {room['area']:.2f} m2  \nPerimetro: {room['perimetro']:.2f} m"
                )


def renderizar_importacao_dxf() -> list[dict[str, Any]]:
    st.subheader('1) Importacao DXF e selecao manual dos comodos')
    st.caption('Envie o DXF e abra o editor para desenhar cada comodo diretamente sobre a planta.')

    rooms_key = _session_key('rooms')
    endpoints_key = _session_key('endpoints')
    segments_key = _session_key('segments')
    open_key = _session_key('open_editor')
    canvas_key = _session_key('canvas_state')
    view_key = _session_key('view_state')
    pending_key = _session_key('pending_room')
    pending_name_key = _session_key('pending_room_name')

    if rooms_key not in st.session_state:
        st.session_state[rooms_key] = []
    if open_key not in st.session_state:
        st.session_state[open_key] = False

    uploaded_file = st.file_uploader(
        'Adicionar arquivo DXF',
        type=['dxf'],
        key=_session_key('upload'),
    )

    if uploaded_file is None:
        return st.session_state[rooms_key]

    file_bytes = uploaded_file.getvalue()
    file_hash = _hash_bytes(file_bytes)

    if st.session_state.get(_session_key('file_hash')) != file_hash:
        try:
            segments, endpoints = _load_dxf_payload(file_bytes)
        except Exception as exc:
            st.error(f'Falha ao ler o DXF: {exc}')
            return st.session_state[rooms_key]

        st.session_state[_session_key('file_hash')] = file_hash
        st.session_state[segments_key] = segments
        st.session_state[endpoints_key] = endpoints
        st.session_state[view_key] = _default_view_state(endpoints)
        st.session_state[rooms_key] = []
        st.session_state[canvas_key] = None
        st.session_state[pending_key] = None
        st.session_state[pending_name_key] = ''
        st.session_state[open_key] = False

    segments = st.session_state.get(segments_key)
    endpoints = st.session_state.get(endpoints_key)
    rooms = st.session_state[rooms_key]

    if not segments or not endpoints:
        st.error('Nao foi possivel montar a visualizacao do DXF.')
        return rooms

    action_col1, action_col2 = st.columns([1.4, 2.2])
    with action_col1:
        if st.button('Abrir editor do projeto em tela ampliada', use_container_width=True):
            st.session_state[open_key] = True
            st.rerun()
    with action_col2:
        st.info('No editor, ajuste a vista, desenhe o c?modo, salve e informe o nome na hora.')

    @st.dialog('Editor DXF', width='large')
    def _editor_dialog() -> None:
        _render_dxf_editor_panel(
            segments=segments,
            endpoints=endpoints,
            rooms_key=rooms_key,
            canvas_key=canvas_key,
            open_key=open_key,
            view_key=view_key,
            pending_key=pending_key,
            pending_name_key=pending_name_key,
            dialog_mode=True,
        )

    if st.session_state.get(open_key):
        _editor_dialog()

    st.markdown('**Comodos importados**')
    if not rooms:
        st.warning('Abra o editor ampliado, desenhe e salve ao menos um comodo no DXF para continuar com o formulario abaixo.')
    else:
        resumo_cols = st.columns(min(3, max(1, len(rooms))))
        for idx, room in enumerate(rooms):
            with resumo_cols[idx % len(resumo_cols)]:
                st.metric(room['nome'], f"{room['area']:.2f} m2", f"Perimetro {room['perimetro']:.2f} m")

    return rooms


CATEGORIAS_DEMANDA = {
    "b": "Chuveiro, torneira eletrica, aquecedor de passagem, ferro eletrico",
    "c": "Boiler / aquecedor central",
    "d": "Secadora, forno eletrico, lava-loucas, micro-ondas",
    "e": "Fogao eletrico",
    "f": "Ar-condicionado tipo janela",
    "g": "Motor / maquina de solda a motor",
    "h": "Equipamento especial",
    "i": "Hidromassagem",
    "x": "Sem categoria GED-13 (considerar FD = 1)",
}

PADRAO_ENTRADA_TABELA_1A = {
    "A1": {
        "tipo_caixa": "II",
        "disjuntor_a": 32,
        "medida_eletroduto": "32 (1)",
    },
    "A2": {
        "tipo_caixa": "II",
        "disjuntor_a": 63,
        "medida_eletroduto": "32 (1)",
    },
    "B1": {
        "tipo_caixa": "II",
        "disjuntor_a": 63,
        "medida_eletroduto": "40 (1 1/4)",
    },
    "B2": {
        "tipo_caixa": "II",
        "disjuntor_a": 80,
        "medida_eletroduto": "40 (1 1/4)",
    },
    "C1": {
        "tipo_caixa": "III",
        "disjuntor_a": 63,
        "medida_eletroduto": "40 (1 1/4)",
    },
    "C2": {
        "tipo_caixa": "III",
        "disjuntor_a": 80,
        "medida_eletroduto": "40 (1 1/4)",
    },
    "C3": {
        "tipo_caixa": "III",
        "disjuntor_a": 100,
        "medida_eletroduto": "40 (1 1/4)",
    },
    "C4": {
        "tipo_caixa": "III",
        "disjuntor_a": 125,
        "medida_eletroduto": "50 (1 1/2)",
    },
    "C5": {
        "tipo_caixa": "H",
        "disjuntor_a": 150,
        "medida_eletroduto": "50 (1 1/2)",
    },
    "C6": {
        "tipo_caixa": "H",
        "disjuntor_a": 200,
        "medida_eletroduto": "60 (2)",
    },
}


# -------------------------------
# FORMATACAO
# -------------------------------
def formatar_numero_br(valor: float, casas: int = 1) -> str:
    texto = f"{valor:,.{casas}f}"
    return texto.replace(",", "X").replace(".", ",").replace("X", ".")


# -------------------------------
# NORMALIZACAO
# -------------------------------
def normalizar_ambiente(ambiente: str) -> str:
    amb = ambiente.strip().lower()
    trocas = {
        "a": ["a", "á", "à", "ã", "â"],
        "e": ["e", "é", "ê"],
        "i": ["i", "í"],
        "o": ["o", "ó", "ô", "õ"],
        "u": ["u", "ú"],
        "c": ["c", "ç"],
    }
    for destino, origens in trocas.items():
        for origem in origens[1:]:
            amb = amb.replace(origem, destino)

    aliases = {
        "suite": "suite",
        "suíte": "suite",
        "dormitorio": "quarto",
        "dormitório": "quarto",
        "sala de estar": "sala",
        "sala de jantar": "sala",
        "area de servico": "area_servico",
        "área de servico": "area_servico",
        "area de serviço": "area_servico",
        "área de serviço": "area_servico",
        "copa-cozinha": "copa_cozinha",
        "copa cozinha": "copa_cozinha",
        "bwc": "banheiro",
        "wc": "banheiro",
        "hall de escadaria": "hall",
        "hall escadaria": "hall",
        "casa de maquinas": "casa_maquinas",
        "casa de máquinas": "casa_maquinas",
        "sala de bombas": "sala_bombas",
    }
    return aliases.get(amb, amb)


# -------------------------------
# NBR 5410 - ILUMINACAO
# -------------------------------
def calcular_iluminacao(area: float) -> Dict[str, int]:
    pontos_minimos = 1

    if area <= 6:
        potencia_va = 100
    else:
        acrescimos = math.floor((area - 6) / 4)
        potencia_va = 100 + acrescimos * 60

    return {
        "pontos_minimos": pontos_minimos,
        "potencia_va": potencia_va,
    }


# -------------------------------
# NBR 5410 - TUG
# -------------------------------
def calcular_tug(
    area: float,
    perimetro: float,
    ambiente: str,
    bancadas_validas: int = 0,
) -> Dict[str, object]:
    amb = normalizar_ambiente(ambiente)

    if amb == "banheiro":
        pontos = 1
        potencias = [600]

    elif amb in {"cozinha", "copa", "copa_cozinha", "area_servico", "lavanderia"}:
        pontos_perimetro = math.ceil(perimetro / 3.5)
        pontos = max(pontos_perimetro, bancadas_validas)
        potencias = [600] * min(pontos, 3) + [100] * max(0, pontos - 3)

    elif amb in {"garagem", "varanda", "sotao", "subsolo", "hall", "casa_maquinas", "sala_bombas", "barrilete"}:
        pontos = 1
        potencias = [100]

    else:
        if area <= 6:
            pontos = 1
        else:
            pontos = math.ceil(perimetro / 5)
        potencias = [100] * pontos

    return {
        "pontos": pontos,
        "potencias_va": potencias,
        "potencia_total_va": sum(potencias),
    }


# -------------------------------
# NBR 5410 - TUE
# -------------------------------
def calcular_tue(equipamentos: List[Dict[str, object]]) -> Dict[str, object]:
    total_w = 0.0
    descricoes = []

    for eq in equipamentos:
        nome = str(eq.get("nome", "")).strip() or "Equipamento"
        potencia = float(eq.get("potencia_w", 0) or 0)
        total_w += potencia
        descricoes.append(f"{nome} ({potencia:.0f} W)")

    return {
        "descricao": " / ".join(descricoes) if descricoes else "-",
        "potencia_total_w": total_w,
    }


# -------------------------------
# TUG - FORMATACAO DO PONTO
# -------------------------------
def formatar_sponto_tug(potencias: List[int]) -> str:
    if not potencias:
        return "-"

    contagem = Counter(potencias)
    partes = []
    for va in sorted(contagem.keys(), reverse=True):
        partes.append(f"{contagem[va]} de {va}")
    return " e ".join(partes)


# -------------------------------
# CPFL / GED-13 - DEMANDA (versao simplificada didatica)
# -------------------------------
def fd_iluminacao_tug(carga_kw: float) -> float:
    if carga_kw <= 1:
        return 0.86
    if carga_kw <= 2:
        return 0.75
    if carga_kw <= 3:
        return 0.66
    if carga_kw <= 4:
        return 0.59
    if carga_kw <= 5:
        return 0.52
    if carga_kw <= 6:
        return 0.45
    if carga_kw <= 7:
        return 0.40
    if carga_kw <= 8:
        return 0.35
    if carga_kw <= 9:
        return 0.31
    if carga_kw <= 10:
        return 0.27
    return 0.24


FD_TABELA_B = {
    1: 1.00,
    2: 1.00,
    3: 0.84,
    4: 0.76,
    5: 0.70,
    6: 0.65,
    7: 0.60,
    8: 0.57,
    9: 0.54,
    10: 0.52,
    11: 0.49,
    12: 0.48,
    13: 0.46,
    14: 0.45,
    15: 0.44,
    16: 0.43,
    17: 0.42,
    18: 0.41,
    19: 0.40,
    20: 0.40,
    21: 0.39,
    22: 0.39,
    23: 0.39,
    24: 0.38,
    25: 0.38,
}


def fd_categoria_b(qtd: int) -> float:
    return FD_TABELA_B.get(qtd, 0.38)



def fd_categoria_c(qtd: int) -> float:
    if qtd <= 1:
        return 1.00
    if qtd == 2:
        return 0.72
    return 0.62



def fd_categoria_d(qtd: int) -> float:
    if qtd <= 1:
        return 1.00
    if 2 <= qtd <= 4:
        return 0.70
    if 5 <= qtd <= 6:
        return 0.60
    return 0.50



def fd_categoria_e(qtd: int) -> float:
    tabela = {
        1: 1.00,
        2: 0.60,
        3: 0.48,
        4: 0.40,
        5: 0.37,
        6: 0.35,
        7: 0.33,
        8: 0.32,
        9: 0.31,
    }
    if qtd in tabela:
        return tabela[qtd]
    if 10 <= qtd <= 11:
        return 0.30
    if 12 <= qtd <= 15:
        return 0.28
    return 0.26



def fd_categoria_f_residencial(qtd: int) -> float:
    return 1.00



def demanda_maiores_primeiro(potencias_w: List[float], fatores_maiores: List[float], fator_demais: float) -> float:
    if not potencias_w:
        return 0.0

    potencias = sorted((float(p) for p in potencias_w), reverse=True)
    demanda = 0.0

    for i, potencia in enumerate(potencias):
        if i < len(fatores_maiores):
            demanda += potencia * fatores_maiores[i]
        else:
            demanda += potencia * fator_demais

    return demanda



def calcular_demanda_cpfl_simplificada(
    carga_iluminacao_va: float,
    carga_tug_va: float,
    equipamentos_tue: List[Dict[str, object]],
) -> Tuple[pd.DataFrame, float]:
    linhas = []

    carga_a_kw = (carga_iluminacao_va + carga_tug_va) / 1000
    fd_a = fd_iluminacao_tug(carga_a_kw)
    demanda_a_w = (carga_iluminacao_va + carga_tug_va) * fd_a
    linhas.append(
        {
            "Categoria": "a",
            "Descricao": "Iluminacao + TUG",
            "Carga instalada": f"{carga_iluminacao_va + carga_tug_va:.0f} VA",
            "FD": fd_a,
            "Demanda": f"{demanda_a_w:.0f} W",
            "Equipamentos": "-",
        }
    )

    grupos = defaultdict(list)
    for eq in equipamentos_tue:
        categoria = str(eq.get("categoria_demanda", "x"))
        grupos[categoria].append(eq)

    for categoria, itens in grupos.items():
        potencias = [float(item.get("potencia_w", 0) or 0) for item in itens]
        nomes = " / ".join(str(item.get("nome", "Equipamento")) for item in itens)
        qtd = len(itens)

        if categoria == "b":
            fd = fd_categoria_b(qtd)
            demanda_w = sum(potencias) * fd
            desc = "Chuveiro / torneira / aquecedor de passagem / ferro"
        elif categoria == "c":
            fd = fd_categoria_c(qtd)
            demanda_w = sum(potencias) * fd
            desc = "Boiler / aquecedor central"
        elif categoria == "d":
            fd = fd_categoria_d(qtd)
            demanda_w = sum(potencias) * fd
            desc = "Secadora / forno eletrico / lava-loucas / micro-ondas"
        elif categoria == "e":
            fd = fd_categoria_e(qtd)
            demanda_w = sum(potencias) * fd
            desc = "Fogao eletrico"
        elif categoria == "f":
            fd = fd_categoria_f_residencial(qtd)
            demanda_w = sum(potencias) * fd
            desc = "Ar-condicionado tipo janela (uso residencial)"
        elif categoria == "g":
            fd = "maiores"
            demanda_w = demanda_maiores_primeiro(potencias, [1.00, 0.90, 0.80, 0.80, 0.80], 0.70)
            desc = "Motores / solda a motor"
        elif categoria == "h":
            fd = 1.00
            demanda_w = sum(potencias)
            desc = "Equipamentos especiais (simplificado)"
        elif categoria == "i":
            fd = "maiores"
            demanda_w = demanda_maiores_primeiro(potencias, [1.00, 0.90, 0.80, 0.80, 0.80], 0.70)
            desc = "Hidromassagem"
        else:
            fd = 1.00
            demanda_w = sum(potencias)
            desc = "Sem categoria GED-13"

        linhas.append(
            {
                "Categoria": categoria,
                "Descricao": desc,
                "Carga instalada": f"{sum(potencias):.0f} W",
                "FD": fd,
                "Demanda": f"{demanda_w:.0f} W",
                "Equipamentos": nomes,
            }
        )

    df = pd.DataFrame(linhas)
    total_demanda_w = float(demanda_a_w)
    for _, row in df.iloc[1:].iterrows():
        total_demanda_w += float(str(row["Demanda"]).replace(" W", ""))

    return df, total_demanda_w


# -------------------------------
# PADRAO DE ENTRADA - TABELA 1A GED-13
# -------------------------------
def determinar_categoria_padrao_entrada(carga_instalada_w: float, demanda_total_w: float) -> Tuple[str, str]:
    carga_instalada_kw = carga_instalada_w / 1000
    demanda_total_kva = demanda_total_w / 1000

    if carga_instalada_kw <= 6:
        return "A1", "carga instalada"
    if carga_instalada_kw <= 12:
        return "A2", "carga instalada"
    if carga_instalada_kw <= 18:
        return "B1", "carga instalada"
    if carga_instalada_kw <= 25:
        return "B2", "carga instalada"

    if demanda_total_kva <= 23:
        return "C1", "demanda"
    if demanda_total_kva <= 30:
        return "C2", "demanda"
    if demanda_total_kva <= 38:
        return "C3", "demanda"
    if demanda_total_kva <= 47:
        return "C4", "demanda"
    if demanda_total_kva <= 57:
        return "C5", "demanda"
    if demanda_total_kva <= 76:
        return "C6", "demanda"

    return "CONSULTAR GED-13", "consultar GED-13"



def resolver_fase_padrao(categoria: str, fase_escolhida: str) -> str:
    if fase_escolhida != "Automatico":
        return fase_escolhida

    if categoria in {"A1", "A2", "B1", "B2"}:
        return "Monofasico"
    if categoria in {"C1", "C2", "C3", "C4", "C5", "C6"}:
        return "Trifasico"
    return "Consultar GED-13"



def calcular_padrao_entrada(
    carga_instalada_w: float,
    demanda_total_w: float,
    fase_escolhida: str,
) -> Dict[str, str]:
    categoria, criterio = determinar_categoria_padrao_entrada(carga_instalada_w, demanda_total_w)
    fase = resolver_fase_padrao(categoria, fase_escolhida)
    caracteristicas = PADRAO_ENTRADA_TABELA_1A.get(categoria, {})

    return {
        "Fase": fase,
        "Categoria": categoria,
        "Carga Instalada": f"{formatar_numero_br(carga_instalada_w, 1)} W",
        "Tipo de Caixa": caracteristicas.get("tipo_caixa", "Consultar GED-13"),
        "Disjuntor": str(caracteristicas.get("disjuntor_a", "Consultar GED-13")),
        "Medida do Eletroduto": caracteristicas.get("medida_eletroduto", "Consultar GED-13"),
        "Criterio da Categoria": criterio,
        "Demanda Total Simplificada": f"{formatar_numero_br(demanda_total_w, 1)} W",
    }


# -------------------------------
# EXPORTACAO XLSX
# -------------------------------
def ajustar_largura_colunas(worksheet) -> None:
    for coluna in worksheet.columns:
        maximo = 0
        indice_coluna = coluna[0].column
        letra = get_column_letter(indice_coluna)
        for celula in coluna:
            valor = "" if celula.value is None else str(celula.value)
            if len(valor) > maximo:
                maximo = len(valor)
        worksheet.column_dimensions[letra].width = min(maximo + 3, 60)



def aplicar_estilo_planilha(writer: pd.ExcelWriter) -> None:
    workbook = writer.book
    cabecalho_fill = PatternFill(fill_type="solid", fgColor="1F4E78")
    cabecalho_font = Font(color="FFFFFF", bold=True)

    for worksheet in workbook.worksheets:
        if worksheet.max_row >= 1:
            for celula in worksheet[1]:
                celula.fill = cabecalho_fill
                celula.font = cabecalho_font
        worksheet.freeze_panes = "A2"
        ajustar_largura_colunas(worksheet)



def gerar_excel_bytes(
    df_resultados: pd.DataFrame,
    df_demanda: pd.DataFrame,
    df_padrao: pd.DataFrame,
    df_resumo: pd.DataFrame,
) -> bytes:
    buffer = io.BytesIO()

    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        df_resultados.to_excel(writer, sheet_name="Tabela de Cargas", index=False)
        df_demanda.to_excel(writer, sheet_name="Demanda CPFL", index=False)
        df_padrao.to_excel(writer, sheet_name="Padrao de Entrada", index=False)
        df_resumo.to_excel(writer, sheet_name="Resumo", index=False)
        aplicar_estilo_planilha(writer)

    buffer.seek(0)
    return buffer.getvalue()


ARQUIVO_SECRETS = Path(__file__).with_name("secrets.toml")
ARQUIVO_OAUTH_ESTADO = Path(__file__).with_name("oauth_state.json")


# -------------------------------
# AUTENTICACAO
# -------------------------------
def carregar_config_autenticacao() -> Dict[str, object]:
    config: Dict[str, object] = {}

    try:
        config.update(dict(st.secrets))
    except Exception:
        pass

    if ARQUIVO_SECRETS.exists() and tomllib is not None:
        with ARQUIVO_SECRETS.open("rb") as arquivo:
            config_arquivo = tomllib.load(arquivo)
        for chave, valor in config_arquivo.items():
            config[chave] = valor

    return config


def obter_query_params() -> Dict[str, str]:
    if hasattr(st, "query_params"):
        params = {}
        for chave in st.query_params.keys():
            valor = st.query_params[chave]
            if isinstance(valor, list):
                params[chave] = valor[0] if valor else ""
            else:
                params[chave] = valor
        return params

    params_legacy = st.experimental_get_query_params()
    return {chave: valores[0] for chave, valores in params_legacy.items() if valores}


def limpar_query_params() -> None:
    if hasattr(st, "query_params"):
        for chave in list(st.query_params.keys()):
            del st.query_params[chave]
    else:
        st.experimental_set_query_params()


def salvar_estado_oauth(state: str, provider: str) -> None:
    dados = {"state": state, "provider": provider}
    ARQUIVO_OAUTH_ESTADO.write_text(json.dumps(dados), encoding="utf-8")


def carregar_estado_oauth() -> Dict[str, str]:
    if not ARQUIVO_OAUTH_ESTADO.exists():
        return {}

    try:
        conteudo = ARQUIVO_OAUTH_ESTADO.read_text(encoding="utf-8")
        dados = json.loads(conteudo)
        if isinstance(dados, dict):
            return {str(chave): str(valor) for chave, valor in dados.items()}
    except Exception:
        return {}

    return {}


def limpar_estado_oauth() -> None:
    try:
        if ARQUIVO_OAUTH_ESTADO.exists():
            ARQUIVO_OAUTH_ESTADO.unlink()
    except Exception:
        pass


def carregar_usuarios_locais(config: Dict[str, object]) -> List[Dict[str, str]]:
    usuarios: List[Dict[str, str]] = []

    json_usuarios = config.get("LOCAL_USERS_JSON")
    if json_usuarios:
        try:
            dados = json.loads(str(json_usuarios))
            if isinstance(dados, list):
                for item in dados:
                    if isinstance(item, dict):
                        usuarios.append({str(k): str(v) for k, v in item.items()})
        except json.JSONDecodeError:
            pass

    email = str(config.get("LOCAL_LOGIN_EMAIL", "")).strip()
    senha = str(config.get("LOCAL_LOGIN_PASSWORD", "")).strip()
    nome = str(config.get("LOCAL_LOGIN_NAME", "Usuario local")).strip() or "Usuario local"
    if email and senha:
        usuarios.append({
            "email": email,
            "password": senha,
            "name": nome,
        })

    return usuarios


def autenticar_login_local(email: str, senha: str, usuarios: List[Dict[str, str]]) -> Dict[str, str] | None:
    email_normalizado = email.strip().lower()

    for usuario in usuarios:
        email_usuario = str(usuario.get("email", "")).strip().lower()
        senha_usuario = str(usuario.get("password", "")).strip()
        if email_usuario == email_normalizado and senha_usuario == senha:
            return {
                "provider": "local",
                "email": email_usuario,
                "name": str(usuario.get("name", email_usuario)).strip() or email_usuario,
            }

    return None


def obter_config_oauth(provider: str, config: Dict[str, object]) -> Dict[str, object] | None:
    if provider == "google":
        client_id = str(config.get("GOOGLE_CLIENT_ID", "")).strip()
        client_secret = str(config.get("GOOGLE_CLIENT_SECRET", "")).strip()
        redirect_uri = str(config.get("GOOGLE_REDIRECT_URI", "")).strip()
        if not client_id or not client_secret or not redirect_uri or "..." in redirect_uri:
            return None
        return {
            "provider": "google",
            "label": "Google",
            "client_id": client_id,
            "client_secret": client_secret,
            "redirect_uri": redirect_uri,
            "authorize_url": "https://accounts.google.com/o/oauth2/v2/auth",
            "token_url": "https://oauth2.googleapis.com/token",
            "userinfo_url": "https://openidconnect.googleapis.com/v1/userinfo",
            "scope": "openid email profile",
            "extra_params": {"prompt": "select_account"},
        }


    return None


def montar_url_autorizacao(provider: str, config: Dict[str, object]) -> str | None:
    oauth = obter_config_oauth(provider, config)
    if oauth is None:
        return None

    state = f"{provider}:{secrets.token_urlsafe(24)}"
    st.session_state["oauth_state"] = state
    salvar_estado_oauth(state, provider)

    params = {
        "client_id": oauth["client_id"],
        "redirect_uri": oauth["redirect_uri"],
        "response_type": "code",
        "scope": oauth["scope"],
        "state": state,
    }
    params.update(oauth.get("extra_params", {}))

    return f"{oauth['authorize_url']}?{urllib.parse.urlencode(params)}"


def criar_contexto_ssl() -> ssl.SSLContext:
    if certifi is not None:
        return ssl.create_default_context(cafile=certifi.where())
    return ssl.create_default_context()


def requisicao_json(url: str, data: Dict[str, str] | None = None, headers: Dict[str, str] | None = None) -> Dict[str, object]:
    headers = headers or {}
    contexto_ssl = criar_contexto_ssl()

    if data is None:
        request = urllib.request.Request(url, headers=headers)
    else:
        body = urllib.parse.urlencode(data).encode("utf-8")
        headers = {"Content-Type": "application/x-www-form-urlencoded", **headers}
        request = urllib.request.Request(url, data=body, headers=headers)

    with urllib.request.urlopen(request, timeout=20, context=contexto_ssl) as resposta:
        return json.loads(resposta.read().decode("utf-8"))


def trocar_code_por_token(provider: str, code: str, config: Dict[str, object]) -> Dict[str, object]:
    oauth = obter_config_oauth(provider, config)
    if oauth is None:
        raise ValueError(f"Configuracao de {provider} nao encontrada.")

    payload = {
        "code": code,
        "client_id": str(oauth["client_id"]),
        "client_secret": str(oauth["client_secret"]),
        "redirect_uri": str(oauth["redirect_uri"]),
        "grant_type": "authorization_code",
    }
    return requisicao_json(str(oauth["token_url"]), data=payload)


def decodificar_id_token_sem_validacao(id_token: str) -> Dict[str, object]:
    partes = id_token.split(".")
    if len(partes) < 2:
        return {}

    payload = partes[1]
    padding = "=" * (-len(payload) % 4)
    try:
        conteudo = urllib.parse.unquote_to_bytes(payload + padding)
    except Exception:
        conteudo = (payload + padding).encode("utf-8")

    try:
        import base64
        bruto = base64.urlsafe_b64decode(payload + padding)
        return json.loads(bruto.decode("utf-8"))
    except Exception:
        try:
            return json.loads(conteudo.decode("utf-8"))
        except Exception:
            return {}


def obter_usuario_oauth(provider: str, token: Dict[str, object], config: Dict[str, object]) -> Dict[str, str]:
    oauth = obter_config_oauth(provider, config)
    if oauth is None:
        raise ValueError(f"Configuracao de {provider} nao encontrada.")

    access_token = str(token.get("access_token", "")).strip()
    if access_token:
        dados = requisicao_json(
            str(oauth["userinfo_url"]),
            headers={"Authorization": f"Bearer {access_token}"},
        )
        return {
            "provider": provider,
            "email": str(dados.get("email") or dados.get("preferred_username") or "").strip(),
            "name": str(dados.get("name") or dados.get("given_name") or dados.get("preferred_username") or "Usuario").strip(),
        }

    id_token = str(token.get("id_token", "")).strip()
    dados = decodificar_id_token_sem_validacao(id_token)
    return {
        "provider": provider,
        "email": str(dados.get("email") or dados.get("preferred_username") or "").strip(),
        "name": str(dados.get("name") or dados.get("given_name") or dados.get("preferred_username") or "Usuario").strip(),
    }


def processar_callback_oauth(config: Dict[str, object]) -> None:
    params = obter_query_params()
    code = params.get("code", "")
    state = params.get("state", "")
    erro = params.get("error", "")

    if erro:
        st.error(f"Falha no login SSO: {erro}.")
        limpar_query_params()
        return

    if not code or not state:
        return

    estado_salvo = carregar_estado_oauth()
    state_esperado = st.session_state.get("oauth_state", "") or estado_salvo.get("state", "")
    if not state_esperado or state != state_esperado:
        st.error("Nao foi possivel validar a resposta do provedor de login.")
        limpar_query_params()
        limpar_estado_oauth()
        return

    provider = estado_salvo.get("provider", "") or state.split(":", 1)[0]

    try:
        token = trocar_code_por_token(provider, code, config)
        usuario = obter_usuario_oauth(provider, token, config)
    except Exception as exc:
        st.error(f"Falha ao concluir o login com {provider.title()}: {exc}")
        limpar_query_params()
        limpar_estado_oauth()
        return

    if not usuario.get("email"):
        st.error("O provedor retornou o login, mas sem um e-mail utilizavel.")
        limpar_query_params()
        limpar_estado_oauth()
        return

    st.session_state["auth_user"] = usuario
    limpar_query_params()
    limpar_estado_oauth()
    st.rerun()


def usuario_autenticado() -> bool:
    return bool(st.session_state.get("auth_user"))


def sair() -> None:
    st.session_state.pop("auth_user", None)
    st.session_state.pop("oauth_state", None)
    limpar_query_params()
    limpar_estado_oauth()
    st.rerun()


def renderizar_botao_oauth(provider: str, config: Dict[str, object]) -> None:
    oauth = obter_config_oauth(provider, config)
    if oauth is None:
        st.warning(f"Login com {provider.title()} ainda nao esta configurado no secrets.toml.")
        return

    url = montar_url_autorizacao(provider, config)
    if not url:
        st.warning(f"Nao foi possivel montar a URL de login do {provider.title()}.")
        return

    label = f"Entrar com {oauth['label']}"
    if hasattr(st, "link_button"):
        st.link_button(label, url, use_container_width=True)
    else:
        st.markdown(f"[{label}]({url})")

    st.caption(f"Redirect URI configurada: {oauth['redirect_uri']}")


def renderizar_tela_login(config: Dict[str, object]) -> None:
    st.title("Bem-vindo ao Projeto TCC")
    st.caption("Escolha como deseja entrar antes de acessar o sistema de dimensionamento.")

    metodo = st.radio(
        "Como deseja entrar?",
        options=["Login local", "SSO Google"],
        horizontal=True,
    )

    if metodo == "Login local":
        usuarios = carregar_usuarios_locais(config)
        with st.form("login_local"):
            email = st.text_input("E-mail")
            senha = st.text_input("Senha", type="password")
            enviar = st.form_submit_button("Entrar", use_container_width=True)

        if enviar:
            usuario = autenticar_login_local(email, senha, usuarios)
            if usuario is None:
                st.error("E-mail ou senha invalidos.")
            else:
                st.session_state["auth_user"] = usuario
                st.rerun()

        if not usuarios:
            st.info("Configure LOCAL_LOGIN_EMAIL e LOCAL_LOGIN_PASSWORD no secrets.toml para habilitar o login local.")

    else:
        st.write("Use sua conta Google para entrar no sistema.")
        renderizar_botao_oauth("google", config)

    st.stop()


def renderizar_aplicacao_principal() -> None:
    usuario = st.session_state.get("auth_user", {})

    with st.sidebar:
        st.success(f"Conectado como: {usuario.get('name', 'Usuario')}")
        st.caption(f"E-mail: {usuario.get('email', '-')}")
        st.caption(f"Acesso: {str(usuario.get('provider', 'local')).title()}")
        if st.button("Sair", use_container_width=True):
            sair()

    st.title("Sistema de Dimensionamento Eletrico Residencial")
    st.caption(
        "Versao corrigida para previsao minima de cargas pela NBR 5410, resumo simplificado de demanda CPFL/GED-13 "
        "e caracteristicas do padrao de entrada conforme a tabela 1A usada nas aulas."
    )

    with st.sidebar:
        st.header("Projeto")
        nome_projeto = st.text_input("Nome do projeto", value="Meu Projeto")
        responsavel = st.text_input("Responsavel", value="")
        origem_comodos = st.radio(
            "Origem dos comodos",
            options=["DXF", "Manual"],
            index=0,
            help="Em DXF, os comodos sao desenhados sobre a planta e o formulario herda nome, area e perimetro automaticamente.",
        )
        if origem_comodos == "Manual":
            numero_comodos = st.number_input("Quantidade de comodos", min_value=1, max_value=30, value=5, step=1)
        else:
            numero_comodos = 0
            st.caption("A quantidade de comodos sera definida pela selecao feita no DXF.")

        st.header("Padrao de Entrada")
        fase_padrao = st.selectbox(
            "Fase do padrao de entrada",
            options=["Automatico", "Monofasico", "Bifasico", "Trifasico"],
            index=0,
            help="Se ficar em Automatico, o app sugere a fase com base na categoria. Se voce quiser, pode fixar manualmente.",
        )

    comodos_importados = []
    if origem_comodos == "DXF":
        comodos_importados = renderizar_importacao_dxf()
        numero_comodos = len(comodos_importados)
        if numero_comodos == 0:
            st.warning("Desenhe e salve ao menos um comodo no DXF para continuar com o formulario abaixo.")

    st.subheader("2) Cadastro dos comodos" if origem_comodos == "DXF" else "1) Cadastro dos comodos")

    comodos = []
    equipamentos_gerais_demanda = []

    for i in range(int(numero_comodos)):
        comodo_importado = comodos_importados[i] if i < len(comodos_importados) else None
        with st.expander(f"Comodo {i + 1}", expanded=(i == 0)):
            col1, col2, col3 = st.columns(3)
            with col1:
                nome_padrao = comodo_importado["nome"] if comodo_importado else f"Comodo {i + 1}"
                nome_comodo = st.text_input(
                    f"Nome do comodo {i + 1}",
                    key=f"nome_{i}",
                    value=nome_padrao,
                    disabled=bool(comodo_importado),
                )
            with col2:
                tipos_comodo_opcoes = [
                    "quarto",
                    "suite",
                    "sala",
                    "cozinha",
                    "banheiro",
                    "area_servico",
                    "lavanderia",
                    "copa",
                    "copa_cozinha",
                    "escritorio",
                    "circulacao",
                    "corredor",
                    "hall",
                    "lavabo",
                    "closet",
                    "varanda",
                    "garagem",
                    "sotao",
                    "subsolo",
                    "casa_maquinas",
                    "sala_bombas",
                    "barrilete",
                    "outro",
                ]
                tipos_comodo_labels = {
                    "quarto": "Quarto",
                    "suite": "Su?te",
                    "sala": "Sala",
                    "cozinha": "Cozinha",
                    "banheiro": "Banheiro",
                    "area_servico": "?rea de Servi?o",
                    "lavanderia": "Lavanderia",
                    "copa": "Copa",
                    "copa_cozinha": "Copa/Cozinha",
                    "escritorio": "Escrit?rio",
                    "circulacao": "Circula??o",
                    "corredor": "Corredor",
                    "hall": "Hall",
                    "lavabo": "Lavabo",
                    "closet": "Closet",
                    "varanda": "Varanda",
                    "garagem": "Garagem",
                    "sotao": "S?t?o",
                    "subsolo": "Subsolo",
                    "casa_maquinas": "Casa de M?quinas",
                    "sala_bombas": "Sala de Bombas",
                    "barrilete": "Barrilete",
                    "outro": "Outro",
                }
                tipo_padrao = comodo_importado.get("tipo_sugerido", _infer_tipo_comodo(nome_padrao)) if comodo_importado else "quarto"
                if tipo_padrao not in tipos_comodo_opcoes:
                    tipo_padrao = "outro"
                tipo_comodo = st.selectbox(
                    f"Tipo do comodo {i + 1}",
                    options=tipos_comodo_opcoes,
                    index=tipos_comodo_opcoes.index(tipo_padrao),
                    format_func=lambda x: tipos_comodo_labels.get(x, x),
                    key=f"tipo_{i}",
                )
            with col3:
                area = st.number_input(
                    f"Area (m2) - {i + 1}",
                    min_value=0.01,
                    value=float(comodo_importado["area"]) if comodo_importado else 10.0,
                    step=0.1,
                    key=f"area_{i}",
                    disabled=bool(comodo_importado),
                )

            col4, col5 = st.columns(2)
            with col4:
                perimetro = st.number_input(
                    f"Perimetro (m) - {i + 1}",
                    min_value=0.01,
                    value=float(comodo_importado["perimetro"]) if comodo_importado else 12.0,
                    step=0.1,
                    key=f"per_{i}",
                    disabled=bool(comodo_importado),
                )
            with col5:
                bancadas_validas = 0
                if normalizar_ambiente(tipo_comodo) in {"cozinha", "copa", "copa_cozinha", "area_servico", "lavanderia"}:
                    bancadas_validas = st.number_input(
                        f"Qtde. de bancadas >= 0,30 m - {i + 1}",
                        min_value=0,
                        value=0,
                        step=1,
                        key=f"bancadas_{i}",
                    )
                else:
                    st.markdown("**Bancadas:** nao se aplica")

            st.markdown("**TUEs do comodo**")
            qtd_tues = st.number_input(
                f"Quantidade de TUEs - {i + 1}",
                min_value=0,
                max_value=10,
                value=0,
                step=1,
                key=f"qtd_tue_{i}",
            )
            tues = []

            for j in range(int(qtd_tues)):
                c1, c2, c3 = st.columns([2, 1, 2])
                with c1:
                    nome_eq = st.text_input(f"Equipamento {j + 1}", key=f"eq_nome_{i}_{j}", value=f"Equipamento {j + 1}")
                with c2:
                    potencia_w = st.number_input(
                        f"Potencia (W) {j + 1}",
                        min_value=0.0,
                        value=1000.0,
                        step=100.0,
                        key=f"eq_pot_{i}_{j}",
                    )
                with c3:
                    categoria_demanda = st.selectbox(
                        f"Categoria demanda {j + 1}",
                        options=list(CATEGORIAS_DEMANDA.keys()),
                        format_func=lambda x: f"{x}) {CATEGORIAS_DEMANDA[x]}",
                        key=f"eq_cat_{i}_{j}",
                    )

                registro_eq = {
                    "nome": nome_eq,
                    "potencia_w": potencia_w,
                    "categoria_demanda": categoria_demanda,
                    "comodo": nome_comodo,
                }
                tues.append(registro_eq)
                equipamentos_gerais_demanda.append(registro_eq)

            comodos.append(
                {
                    "nome": nome_comodo,
                    "tipo": tipo_comodo,
                    "area": area,
                    "perimetro": perimetro,
                    "bancadas_validas": bancadas_validas,
                    "tues": tues,
                }
            )

    if st.button("Calcular projeto", type="primary"):
        resultados = []
        total_ilum_va = 0.0
        total_tug_va = 0.0
        total_tue_w = 0.0

        for comodo in comodos:
            ilum = calcular_iluminacao(comodo["area"])
            tug = calcular_tug(
                area=comodo["area"],
                perimetro=comodo["perimetro"],
                ambiente=comodo["tipo"],
                bancadas_validas=int(comodo["bancadas_validas"]),
            )
            tue = calcular_tue(comodo["tues"])

            total_ilum_va += ilum["potencia_va"]
            total_tug_va += tug["potencia_total_va"]
            total_tue_w += tue["potencia_total_w"]

            resultados.append(
                {
                    "Comodo": comodo["nome"],
                    "Tipo": comodo["tipo"],
                    "Area (m2)": round(comodo["area"], 2),
                    "Perimetro (m)": round(comodo["perimetro"], 2),
                    "Iluminacao - Pontos min.": ilum["pontos_minimos"],
                    "Iluminacao - Stotal (VA)": ilum["potencia_va"],
                    "TUG - Pontos": tug["pontos"],
                    "TUG - Sponto (VA)": formatar_sponto_tug(tug["potencias_va"]),
                    "TUG - Stotal (VA)": tug["potencia_total_va"],
                    "TUE - Equipamentos": tue["descricao"],
                    "TUE - Ptotal (W)": round(tue["potencia_total_w"], 1),
                }
            )

        df_resultados = pd.DataFrame(resultados)

        st.subheader("2) Tabela de cargas por comodo")
        st.dataframe(df_resultados, use_container_width=True)

        potencia_instalada_total_w = total_ilum_va + total_tug_va + total_tue_w

        st.subheader("3) Totais instalados")
        c1, c2, c3, c4 = st.columns(4)
        c1.metric("Iluminacao total", f"{formatar_numero_br(total_ilum_va, 1)} VA")
        c2.metric("TUG total", f"{formatar_numero_br(total_tug_va, 1)} VA")
        c3.metric("TUE total", f"{formatar_numero_br(total_tue_w, 1)} W")
        c4.metric("Carga instalada total", f"{formatar_numero_br(potencia_instalada_total_w, 1)} W")

        st.info(
            "Observacao: a NBR 5410 define a carga minima de iluminacao em VA para dimensionamento; "
            "isso nao determina automaticamente a quantidade real de lampadas/luminarias do projeto."
        )

        st.subheader("4) Resumo simplificado de demanda CPFL / GED-13")
        df_demanda, total_demanda_w = calcular_demanda_cpfl_simplificada(
            carga_iluminacao_va=total_ilum_va,
            carga_tug_va=total_tug_va,
            equipamentos_tue=equipamentos_gerais_demanda,
        )
        st.dataframe(df_demanda, use_container_width=True)
        st.metric("Demanda total simplificada", f"{formatar_numero_br(total_demanda_w, 1)} W")

        st.subheader("5) Caracteristicas do Padrao de Entrada")
        padrao_entrada = calcular_padrao_entrada(
            carga_instalada_w=potencia_instalada_total_w,
            demanda_total_w=total_demanda_w,
            fase_escolhida=fase_padrao,
        )

        df_padrao = pd.DataFrame(
            {
                "Caracteristica": [
                    "Fase",
                    "Categoria",
                    "Carga Instalada",
                    "Tipo de Caixa",
                    "Disjuntor",
                    "Medida do Eletroduto",
                ],
                "Valor": [
                    padrao_entrada["Fase"],
                    padrao_entrada["Categoria"],
                    padrao_entrada["Carga Instalada"],
                    padrao_entrada["Tipo de Caixa"],
                    padrao_entrada["Disjuntor"],
                    padrao_entrada["Medida do Eletroduto"],
                ],
            }
        )
        st.table(df_padrao)

        st.caption(
            "Categoria sugerida conforme a tabela 1A da GED-13 usada nas aulas. "
            "Se necessario, a fase pode ser ajustada manualmente no menu lateral."
        )

        st.subheader("6) Exportacao")
        df_resumo = pd.DataFrame(
            {
                "Indicador": [
                    "Projeto",
                    "Responsavel",
                    "Quantidade de comodos",
                    "Iluminacao total (VA)",
                    "TUG total (VA)",
                    "TUE total (W)",
                    "Carga instalada total (W)",
                    "Demanda total simplificada (W)",
                    "Fase padrao",
                    "Categoria padrao",
                    "Tipo de caixa",
                    "Disjuntor (A)",
                    "Medida do eletroduto",
                ],
                "Valor": [
                    nome_projeto,
                    responsavel or "-",
                    len(comodos),
                    total_ilum_va,
                    total_tug_va,
                    total_tue_w,
                    potencia_instalada_total_w,
                    total_demanda_w,
                    padrao_entrada["Fase"],
                    padrao_entrada["Categoria"],
                    padrao_entrada["Tipo de Caixa"],
                    padrao_entrada["Disjuntor"],
                    padrao_entrada["Medida do Eletroduto"],
                ],
            }
        )

        excel_bytes = gerar_excel_bytes(
            df_resultados=df_resultados,
            df_demanda=df_demanda,
            df_padrao=df_padrao,
            df_resumo=df_resumo,
        )

        st.download_button(
            label="Baixar relatorio em XLSX",
            data=excel_bytes,
            file_name=f"relatorio_eletrico_{nome_projeto.replace(' ', '_').lower()}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

        st.subheader("7) Conferencia rapida")
        st.markdown(
            f"""
            **Projeto:** {nome_projeto}  
            **Responsavel:** {responsavel or '-'}  
            **Quantidade de comodos:** {len(comodos)}  
            **Soma iluminacao:** {formatar_numero_br(total_ilum_va, 1)} VA  
            **Soma TUG:** {formatar_numero_br(total_tug_va, 1)} VA  
            **Soma TUE:** {formatar_numero_br(total_tue_w, 1)} W  
            **Carga instalada total:** {formatar_numero_br(potencia_instalada_total_w, 1)} W  
            **Demanda simplificada:** {formatar_numero_br(total_demanda_w, 1)} W  
            **Padrao sugerido:** {padrao_entrada['Categoria']} / {padrao_entrada['Fase']} / Caixa {padrao_entrada['Tipo de Caixa']} / Disj. {padrao_entrada['Disjuntor']} A / Eletroduto {padrao_entrada['Medida do Eletroduto']}
            """
        )


def main() -> None:
    config = carregar_config_autenticacao()
    processar_callback_oauth(config)

    if not usuario_autenticado():
        renderizar_tela_login(config)

    renderizar_aplicacao_principal()


main()
