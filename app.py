import csv
import io
import json
import html
import urllib.parse
import uuid
from pathlib import Path

import streamlit as st

try:
    from openpyxl import load_workbook
except Exception:
    load_workbook = None


APP_DIR = Path(__file__).parent
DATA_FILE = APP_DIR / "dados_casamento.json"
DEFAULT_XLSX = Path(r"C:\Users\gedie\OneDrive\Desktop\Pasta1.xlsx")

INITIAL_FAMILIES = [
    ["Sebastião", "Maria", "Elizeu", "Raquel", "Lucas"],
    ["Alessandro", "Jaqueline", "Diogo", "Eliza", "Nelis", "Vilmar"],
    ["Alcebiados", "Zelita"],
    ["Wesli"],
    ["Marcos", "Elci", "Geise"],
    ["Micael", "Manuele", "Elói"],
    ["Dora", "Luiza", "Vitor"],
    ["Valdir", "Roseli", "Erica"],
    ["Elder", "Priscila", "Alice"],
    ["José"],
    ["Amanda", "Alice", "Mirtes", "Gentil"],
    ["Marcolina", "Beto"],
    ["Eduardo", "Mari", "Heitor"],
    ["Olavo", "Elizangela", "Jonatan"],
    ["Filho", "Ana", "Jonatan", "Evilyn"],
    ["Victor", "Fledson", "Neia"],
    ["Eliel", "Monalisa", "Isaque"],
    ["Marcelo", "Juliana", "Antony"],
    ["Adriano", "Manuele"],
    ["Gesiel", "Paula", "Wendril Gabriel"],
]


def family_label(index, names):
    first_name = names[0] if names else f"Família {index}"
    return f"Família {first_name}"


def display_family_name(family):
    family = (family or "Sem família").strip()
    if family.startswith("Familia "):
        return "Família " + family[len("Familia "):]
    if family == "Sem familia":
        return "Sem família"
    return family


def new_guest(name="", family="", favorite=False):
    return {
        "id": str(uuid.uuid4()),
        "name": name,
        "family": family,
        "favorite": favorite,
        "status": "Pendente",
        "side": "Noiva",
    }


def default_data():
    guests = []
    for index, names in enumerate(INITIAL_FAMILIES, start=1):
        family = family_label(index, names)
        guests.extend(new_guest(name, family) for name in names)
    return {"guests": guests}


def normalize_guest(item):
    if "role" in item:
        favorite = item.get("role") in ["Padrinho", "Madrinha"]
    else:
        favorite = bool(item.get("favorite", False))

    return {
        "id": item.get("id", str(uuid.uuid4())),
        "name": item.get("name", ""),
        "family": item.get("family", "Sem família"),
        "favorite": favorite,
        "status": item.get("status", "Pendente"),
        "side": item.get("side", "Noiva"),
    }


def load_data():
    if DATA_FILE.exists():
        try:
            data = json.loads(DATA_FILE.read_text(encoding="utf-8"))
            if "guests" in data:
                return {"guests": [normalize_guest(g) for g in data["guests"]]}
            if "people" in data:
                return {"guests": [normalize_guest(g) for g in data["people"]]}
        except json.JSONDecodeError:
            pass
    return default_data()


def save_data():
    DATA_FILE.write_text(
        json.dumps({"guests": st.session_state.guests}, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )


def import_xlsx(path):
    if load_workbook is None:
        st.error("Instale openpyxl para importar Excel.")
        return

    workbook = load_workbook(path, read_only=True, data_only=True)
    sheet = workbook.active
    guests = []

    for row_number, row in enumerate(sheet.iter_rows(values_only=True), start=1):
        names = []
        favorites = []
        for index, value in enumerate(row):
            if value is None:
                continue
            text = str(value).strip()
            if not text or text.isdigit() or text.lower().startswith("possiveis"):
                continue
            if index == 7:
                favorites.append(text)
            elif index < 7:
                names.append(text)

        if names:
            family = family_label(row_number, names)
            guests.extend(new_guest(name, family) for name in names)
        for name in favorites:
            guests.append(new_guest(name, "Favoritos da planilha", True))

    st.session_state.guests = guests
    save_data()


def update_guest(guest_id, field, value):
    for guest in st.session_state.guests:
        if guest["id"] == guest_id:
            guest[field] = value
            break
    save_data()


def remove_guest(guest_id):
    st.session_state.guests = [
        guest for guest in st.session_state.guests if guest["id"] != guest_id
    ]
    if st.session_state.get("selected_guest_id") == guest_id:
        st.session_state.selected_guest_id = None
    if st.session_state.get("editing_guest_id") == guest_id:
        st.session_state.editing_guest_id = None
    save_data()


def add_guest():
    st.session_state.guests.insert(0, new_guest("", "Sem família"))
    save_data()


def add_guest_from_inputs():
    name = st.session_state.get("new_guest_name", "").strip()
    family = st.session_state.get("new_guest_family", "").strip() or "Sem família"
    side = st.session_state.get("new_guest_side", "Noiva")
    if not name:
        return
    guest = new_guest(name, family)
    guest["side"] = side
    st.session_state.guests.insert(0, guest)
    st.session_state.new_guest_name = ""
    save_data()


def csv_bytes():
    buffer = io.StringIO()
    writer = csv.writer(buffer)
    writer.writerow(["nome", "familia", "favorito"])
    for guest in st.session_state.guests:
        writer.writerow([guest["name"], guest["family"], "sim" if guest["favorite"] else "nao"])
    return buffer.getvalue().encode("utf-8-sig")


def family_summary():
    families = {}
    for guest in st.session_state.guests:
        if not guest["name"].strip():
            continue
        family = display_family_name(guest["family"].strip() or "Sem família")
        families.setdefault(family, {"total": 0, "favorites": 0})
        families[family]["total"] += 1
        families[family]["favorites"] += int(guest["favorite"])
    return families


def stable_key(value):
    return uuid.uuid5(uuid.NAMESPACE_URL, str(value)).hex


def family_matches(guest, family):
    current = display_family_name((guest.get("family") or "Sem família").strip() or "Sem família")
    return current == display_family_name(family)


def rename_family(old_family, new_family):
    new_family = display_family_name((new_family or "").strip() or "Sem família")
    for guest in st.session_state.guests:
        if family_matches(guest, old_family):
            guest["family"] = new_family
    st.session_state.selected_family = new_family
    save_data()


def add_guest_to_family(family, input_key):
    name = st.session_state.get(input_key, "").strip()
    if not name:
        return
    st.session_state.guests.insert(0, new_guest(name, display_family_name(family)))
    st.session_state[input_key] = ""
    save_data()


def render_family_editor(family):
    safe_family = html.escape(display_family_name(family))
    key = stable_key(family)
    members = [guest for guest in st.session_state.guests if family_matches(guest, family)]

    with st.container(border=True):
        st.markdown(
            f"""
            <div class="family-editor-title">Editar {safe_family}</div>
            <div class="family-editor-sub">Altere o nome da família, adicione pessoas ou edite os convidados desse núcleo.</div>
            """,
            unsafe_allow_html=True,
        )

        name_cols = st.columns([3, 1], vertical_alignment="bottom")
        new_family_name = name_cols[0].text_input(
            "Nome da família",
            value=display_family_name(family),
            key=f"family_rename_{key}",
            placeholder="Ex.: Família Costa",
        )
        if name_cols[1].button("Salvar", key=f"save_family_{key}", type="primary", use_container_width=True):
            rename_family(family, new_family_name)
            st.rerun()

        add_key = f"family_add_guest_{key}"
        add_cols = st.columns([3, 1], vertical_alignment="bottom")
        add_cols[0].text_input(
            "Adicionar convidado nesta família",
            key=add_key,
            placeholder="Nome do convidado",
        )
        if add_cols[1].button("Adicionar", key=f"add_to_family_{key}", use_container_width=True):
            add_guest_to_family(family, add_key)
            st.rerun()

        st.markdown(f"<div class='section-kicker'>{len(members)} convidados nesta família</div>", unsafe_allow_html=True)
        for guest in members:
            guest_row(guest, compact=True)

        if st.button("Fechar edição", key=f"close_family_{key}", use_container_width=True):
            st.session_state.selected_family = None
            st.session_state.editing_guest_id = None
            st.session_state.selected_guest_id = None
            st.rerun()


def guest_row(guest, compact=False):
    selected = st.session_state.get("selected_guest_id") == guest["id"]
    editing = st.session_state.get("editing_guest_id") == guest["id"]
    family_name = display_family_name(guest["family"].strip() or "Sem família")

    with st.container(border=True):
        if editing:
            cols = st.columns([0.45, 2.1, 1.7, 1.25, 1.15, 0.45, 0.45], vertical_alignment="bottom")
        elif selected:
            cols = st.columns([2.5, 2, 1.15, 1, 0.45, 0.45], vertical_alignment="center")
        else:
            cols = st.columns([2.5, 2, 1.15, 1, 0.45], vertical_alignment="center")

        if editing:
            star = "★" if guest["favorite"] else "☆"
            if cols[0].button(star, key=f"fav_edit_{guest['id']}", help="Marcar favorito"):
                update_guest(guest["id"], "favorite", not guest["favorite"])
                st.rerun()

            name = cols[1].text_input(
                "Nome",
                value=guest["name"],
                key=f"name_{guest['id']}",
                placeholder="Nome do convidado",
                label_visibility="collapsed",
            )
            family = cols[2].text_input(
                "Família",
                value=display_family_name(guest["family"]),
                key=f"family_{guest['id']}",
                placeholder="Família",
                label_visibility="collapsed",
            )
            status = cols[3].selectbox(
                "Status",
                ["Confirmado", "Pendente"],
                index=["Confirmado", "Pendente"].index(guest.get("status", "Pendente")),
                key=f"status_{guest['id']}",
                label_visibility="collapsed",
            )
            side = cols[4].selectbox(
                "Lado",
                ["Noiva", "Noivo"],
                index=["Noiva", "Noivo"].index(guest.get("side", "Noiva")),
                key=f"side_{guest['id']}",
                label_visibility="collapsed",
            )
            if name != guest["name"]:
                update_guest(guest["id"], "name", name)
            if family != guest["family"]:
                update_guest(guest["id"], "family", family)
            if status != guest.get("status", "Pendente"):
                update_guest(guest["id"], "status", status)
            if side != guest.get("side", "Noiva"):
                update_guest(guest["id"], "side", side)

            if cols[5].button("✓", key=f"done_{guest['id']}", help="Concluir edição"):
                st.session_state.editing_guest_id = None
                st.rerun()
            if cols[6].button("X", key=f"remove_edit_{guest['id']}", help="Remover"):
                remove_guest(guest["id"])
                st.rerun()
        else:
            name = guest["name"].strip() or "Sem nome"
            family = family_name
            if cols[0].button(
                f"○  {name}",
                key=f"select_{guest['id']}",
                help="Mostrar ações",
                use_container_width=True,
            ):
                st.session_state.selected_guest_id = (
                    None if selected else guest["id"]
                )
                st.session_state.editing_guest_id = None
                st.rerun()

            cols[1].markdown(f"<span class='muted-cell'>☷ {family}</span>", unsafe_allow_html=True)
            status = guest.get("status", "Pendente")
            status_class = "status-ok" if status == "Confirmado" else "status-favorite"
            cols[2].markdown(f"<span class='{status_class}'>{status}</span>", unsafe_allow_html=True)
            side_class = "status-bride" if guest.get("side", "Noiva") == "Noiva" else "status-groom"
            cols[3].markdown(f"<span class='{side_class}'>{guest.get('side', 'Noiva')}</span>", unsafe_allow_html=True)

            if selected:
                if cols[4].button("⋯", key=f"edit_{guest['id']}", help="Editar"):
                    st.session_state.editing_guest_id = guest["id"]
                    st.rerun()
                if cols[5].button("🗑", key=f"remove_{guest['id']}", help="Remover"):
                    remove_guest(guest["id"])
                    st.rerun()
            else:
                if cols[4].button("⋯", key=f"menu_{guest['id']}", help="Editar"):
                    st.session_state.selected_guest_id = guest["id"]
                    st.session_state.editing_guest_id = guest["id"]
                    st.rerun()



def render_header(title, subtitle, icon):
    st.markdown(
        f"""
        <div class="app-header">
            <div>
                <h1 class="app-title">{html.escape(title)}</h1>
                <div class="app-subtitle">{html.escape(subtitle)}</div>
            </div>
            <div class="header-icon">{icon}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def render_family_stats(family_count, total_guests):
    st.markdown(
        f"""
        <div class="stats-card">
            <div class="stats-item">
                <div class="stats-icon">👥</div>
                <div>
                    <div class="stats-number">{family_count}</div>
                    <div class="stats-label">famílias</div>
                </div>
            </div>
            <div class="stats-divider"></div>
            <div class="stats-item">
                <div class="stats-icon">♡</div>
                <div>
                    <div class="stats-number">{total_guests}</div>
                    <div class="stats-label">convidados</div>
                </div>
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def render_family_row(family, values):
    safe_family = html.escape(display_family_name(family))
    total = int(values.get("total", 0))
    favoritos = int(values.get("favorites", 0))
    plural = "convidado" if total == 1 else "convidados"
    fav_text = f" · {favoritos} favorito" if favoritos == 1 else (f" · {favoritos} favoritos" if favoritos else "")
    st.markdown(
        f"""
        <div class="family-row family-row-mobile">
            <div class="family-left">
                <div class="family-avatar">👥</div>
                <div class="family-copy">
                    <div class="family-name">{safe_family}</div>
                    <div class="family-sub">{total} {plural}{fav_text}</div>
                </div>
            </div>
            <div class="family-right">
                <div class="pill">{total}</div>
                <div class="chevron">›</div>
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def nav_href(page):
    return "?" + urllib.parse.urlencode({"page": page})


def set_page(page):
    st.session_state.page = page
    st.session_state.selected_guest_id = None
    st.session_state.editing_guest_id = None
    st.query_params["page"] = page


def render_bottom_nav(active_page):
    nav_items = [
        ("Convidados", "👥"),
        ("Famílias", "⌂"),
        ("Padrinhos", "♡"),
    ]
    st.markdown('<div id="mobile-bottom-nav-marker"></div>', unsafe_allow_html=True)
    nav_cols = st.columns(3)
    for col, (label, icon) in zip(nav_cols, nav_items):
        button_type = "primary" if label == active_page else "secondary"
        if col.button(
            f"{icon}\n{label}",
            key=f"mobile_nav_{label}",
            type=button_type,
            use_container_width=True,
        ):
            set_page(label)
            st.rerun()


st.set_page_config(page_title="Lista do casamento", layout="wide")

st.markdown(
    """
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Playfair+Display:wght@700&family=Inter:wght@400;500;600;700&display=swap');
    :root {
        --rose: #c77d75;
        --rose-dark: #b86d66;
        --rose-light: #f7ece9;
        --ink: #171827;
        --muted: #858891;
        --line: #eadfda;
        --card: rgba(255, 255, 255, 0.90);
        --wash: #fbf6f2;
        --shadow: 0 18px 46px rgba(72, 41, 35, 0.08);
    }
    html, body, [data-testid="stAppViewContainer"] {
        background:
            radial-gradient(circle at 16% 5%, rgba(199, 125, 117, 0.12), transparent 32%),
            radial-gradient(circle at 95% 0%, rgba(199, 125, 117, 0.08), transparent 28%),
            linear-gradient(180deg, #fffbf8 0%, #f9f4ef 100%);
    }
    .block-container {
        padding: 2.2rem 2.1rem 5.8rem;
        max-width: 1340px;
    }
    section[data-testid="stSidebar"] {
        background:
            radial-gradient(circle at 45% 10%, rgba(199, 125, 117, 0.1), transparent 32%),
            rgba(255, 251, 248, 0.96);
        border-right: 1px solid var(--line);
    }
    section[data-testid="stSidebar"] .block-container {
        padding: 2.1rem 1.1rem;
    }
    h1, .app-title {
        color: var(--rose);
        font-family: 'Playfair Display', Georgia, serif;
        font-size: clamp(2.45rem, 7vw, 4.35rem);
        line-height: 0.95;
        letter-spacing: 0;
        margin: 0;
    }
    h2, h3, p, label, div, span, input, button, a {
        font-family: 'Inter', system-ui, sans-serif;
    }
    .app-header {
        display: flex;
        align-items: flex-start;
        justify-content: space-between;
        gap: 1rem;
        margin-bottom: 1.05rem;
    }
    .app-subtitle,
    [data-testid="stCaptionContainer"] {
        color: var(--muted);
        font-size: 1.02rem;
        margin-top: 0.55rem;
    }
    .header-icon {
        display: none;
        color: var(--rose);
        font-size: 1.55rem;
        line-height: 1;
        padding-top: 0.4rem;
        opacity: 0.95;
    }
    div[data-testid="stMetric"] {
        background: var(--card);
        border: 1px solid var(--line);
        border-radius: 16px;
        box-shadow: var(--shadow);
        padding: 14px 16px;
    }
    div[data-testid="stMetricValue"] {
        color: var(--ink);
        font-family: 'Playfair Display', Georgia, serif;
    }
    div[data-testid="stVerticalBlockBorderWrapper"] {
        border-color: var(--line);
        border-radius: 16px;
        background: var(--card);
        box-shadow: 0 10px 25px rgba(72, 41, 35, 0.06);
    }
    .stButton button {
        border-radius: 12px;
        min-height: 42px;
        border-color: var(--line);
        color: var(--rose-dark);
        background: rgba(255, 255, 255, 0.9);
        font-weight: 600;
    }
    .stButton button[kind="primary"], .stDownloadButton button {
        background: linear-gradient(135deg, #dca8a0, #c8746f);
        border: 0;
        color: #fff;
    }
    div[data-testid="stVerticalBlockBorderWrapper"] .stButton button {
        justify-content: flex-start;
        background: transparent;
        border: 0;
        color: var(--ink);
        font-weight: 600;
        box-shadow: none;
    }
    div[data-testid="stVerticalBlockBorderWrapper"] .stButton button:hover {
        background: rgba(199, 125, 117, 0.08);
        color: var(--rose-dark);
    }
    .muted-cell {
        display: inline-flex;
        align-items: center;
        min-height: 42px;
        color: var(--muted);
        font-size: 0.95rem;
    }
    .qty-pill {
        display: inline-flex;
        align-items: center;
        justify-content: center;
        min-width: 58px;
        min-height: 34px;
        border: 1px solid var(--line);
        border-radius: 999px;
        color: var(--ink);
        background: rgba(255, 255, 255, 0.74);
        font-weight: 600;
    }
    .status-ok,
    .status-favorite,
    .status-bride,
    .status-groom {
        display: inline-flex;
        align-items: center;
        justify-content: center;
        min-height: 30px;
        padding: 0 0.8rem;
        border-radius: 10px;
        font-weight: 600;
        font-size: 0.86rem;
    }
    .status-ok {
        color: #427247;
        background: #eaf4e8;
        border: 1px solid #d7ead3;
    }
    .status-favorite {
        color: #b46b1f;
        background: #fff0db;
        border: 1px solid #f3dec0;
    }
    .status-bride {
        color: #9f5f87;
        background: #fbecf5;
        border: 1px solid #efd4e4;
    }
    .status-groom {
        color: #4d638d;
        background: #edf2fb;
        border: 1px solid #d7e0f2;
    }
    .table-card {
        border: 1px solid var(--line);
        border-radius: 18px;
        background: rgba(255, 255, 255, 0.78);
        box-shadow: var(--shadow);
        padding: 0.9rem;
        margin-top: 1rem;
    }
    .table-header {
        color: var(--muted);
        font-size: 0.78rem;
        font-weight: 700;
        letter-spacing: .04em;
        text-transform: uppercase;
        padding: 0.55rem 0.5rem 0.8rem;
        border-bottom: 1px solid var(--line);
        margin-bottom: 0.55rem;
    }
    input {
        border-radius: 14px !important;
    }
    .hero-card {
        border: 1px solid var(--line);
        background: var(--card);
        border-radius: 18px;
        padding: 1rem;
        box-shadow: var(--shadow);
        margin: 1.15rem 0 1.05rem;
    }
    .soft-card, .stats-card {
        border: 1px solid var(--line);
        background: var(--card);
        border-radius: 18px;
        padding: 1.1rem 1.2rem;
        box-shadow: var(--shadow);
        margin: 1.1rem 0 1.2rem;
    }
    .stats-card {
        display: flex;
        align-items: center;
        justify-content: space-around;
        gap: 1rem;
        min-height: 112px;
    }
    .stats-item {
        display: flex;
        align-items: center;
        gap: 1rem;
        min-width: 0;
    }
    .stats-icon, .family-avatar {
        width: 58px;
        height: 58px;
        border-radius: 50%;
        display: flex;
        align-items: center;
        justify-content: center;
        background: var(--rose-light);
        color: var(--rose);
        font-size: 1.45rem;
        flex: 0 0 auto;
    }
    .stats-number {
        color: var(--ink);
        font-family: 'Playfair Display', Georgia, serif;
        font-size: 2.25rem;
        font-weight: 700;
        line-height: 1;
    }
    .stats-label {
        color: var(--muted);
        font-size: 1rem;
        margin-top: 0.35rem;
    }
    .stats-divider {
        width: 1px;
        align-self: stretch;
        background: var(--line);
    }
    .family-row {
        display: flex;
        align-items: center;
        justify-content: space-between;
        border: 1px solid var(--line);
        background: var(--card);
        border-radius: 18px;
        padding: 1rem 1.15rem;
        margin: 0.78rem 0;
        box-shadow: 0 10px 25px rgba(72, 41, 35, 0.05);
    }
    .family-left, .family-right {
        display: flex;
        align-items: center;
        gap: 1rem;
        min-width: 0;
    }
    .family-copy {
        min-width: 0;
    }
    .family-name {
        color: var(--ink);
        font-family: 'Playfair Display', Georgia, serif;
        font-size: 1.35rem;
        font-weight: 700;
        white-space: nowrap;
        overflow: hidden;
        text-overflow: ellipsis;
    }
    .family-sub {
        color: var(--muted);
        margin-top: 0.2rem;
    }
    .pill {
        background: linear-gradient(135deg, #dca8a0, #c8746f);
        color: white;
        border-radius: 13px;
        padding: 0.48rem 0.75rem;
        font-weight: 700;
        min-width: 2.65rem;
        text-align: center;
        box-shadow: 0 8px 16px rgba(199, 125, 117, .22);
    }
    .chevron {
        color: var(--muted);
        font-size: 2rem;
        line-height: 1;
    }
    .bottom-nav {
        display: none;
    }
    .section-kicker {
        color: var(--muted);
        font-weight: 600;
        margin: 0.35rem 0 0.8rem;
    }
    .search-helper {
        display: flex;
        align-items: center;
        justify-content: center;
        height: 44px;
        border: 1px solid var(--line);
        border-radius: 14px;
        background: rgba(255, 255, 255, 0.72);
        color: var(--rose-dark);
        font-size: 1.25rem;
    }
    .family-editor-title {
        color: var(--ink);
        font-family: 'Playfair Display', Georgia, serif;
        font-size: 1.65rem;
        font-weight: 700;
        margin-bottom: 0.25rem;
    }
    .family-editor-sub {
        color: var(--muted);
        font-size: 0.95rem;
        margin-bottom: 0.9rem;
    }
    #mobile-bottom-nav-marker {
        display: none;
    }

    @media (max-width: 720px) {
        .block-container {
            padding: 1.55rem 1.05rem 6.8rem;
            max-width: 480px;
        }
        section[data-testid="stSidebar"] {
            display: none;
        }
        .app-header {
            margin: 0.25rem 0 1.25rem;
        }
        .app-title {
            font-size: 3.05rem;
            line-height: 0.95;
        }
        .app-subtitle {
            font-size: 1rem;
            line-height: 1.35;
            margin-top: 0.55rem;
        }
        .header-icon {
            display: block;
        }
        .hero-card {
            padding: 0.9rem;
            margin-top: 1rem;
            border-radius: 18px;
        }
        .soft-card, .stats-card {
            border-radius: 20px;
            padding: 1rem 1.15rem;
            margin: 1.25rem 0 1rem;
        }
        .stats-card {
            min-height: 104px;
            justify-content: space-between;
        }
        .stats-item {
            gap: 0.9rem;
            flex: 1 1 0;
        }
        .stats-icon, .family-avatar {
            width: 50px;
            height: 50px;
            font-size: 1.25rem;
        }
        .stats-number {
            font-size: 2.05rem;
        }
        .stats-label {
            font-size: .95rem;
        }
        .table-card {
            border: 0;
            background: transparent;
            box-shadow: none;
            padding: 0;
        }
        .table-header {
            display: none;
        }
        div[data-testid="stVerticalBlockBorderWrapper"] {
            background: rgba(255, 255, 255, 0.92);
            border-radius: 18px;
            margin: 0.72rem 0;
            padding: 0.15rem;
            box-shadow: 0 10px 24px rgba(72, 41, 35, 0.06);
        }
        div[data-testid="stVerticalBlockBorderWrapper"] div[data-testid="stHorizontalBlock"] {
            align-items: center;
            gap: 0.35rem;
        }
        div[data-testid="stVerticalBlockBorderWrapper"] .stButton button {
            min-height: 44px;
            border-radius: 14px;
        }
        .muted-cell {
            min-height: 28px;
            padding-left: 0.35rem;
            font-size: 0.9rem;
        }
        .status-ok,
        .status-favorite,
        .status-bride,
        .status-groom {
            min-height: 28px;
            padding: 0 0.65rem;
            font-size: 0.76rem;
        }
        .family-row {
            border-radius: 18px;
            padding: 1rem;
            margin: 0.72rem 0;
        }
        .family-left {
            gap: 0.95rem;
            overflow: hidden;
        }
        .family-right {
            gap: 0.75rem;
            flex: 0 0 auto;
        }
        .family-name {
            font-size: 1.25rem;
        }
        .family-sub {
            font-size: 0.95rem;
        }
        .pill {
            border-radius: 12px;
            min-width: 2.45rem;
            padding: 0.46rem 0.68rem;
        }
        .chevron {
            font-size: 1.8rem;
        }
        /* Menu inferior mobile feito com botões Streamlit.
           Assim a navegação troca a tela dentro do app e não abre outra aba/navegador. */
        div[data-testid="stElementContainer"]:has(#mobile-bottom-nav-marker) + div[data-testid="stHorizontalBlock"],
        .element-container:has(#mobile-bottom-nav-marker) + div[data-testid="stHorizontalBlock"] {
            position: fixed;
            left: 0;
            right: 0;
            bottom: 0;
            display: grid !important;
            grid-template-columns: repeat(3, 1fr);
            gap: 0;
            background: rgba(255,255,255,0.92);
            backdrop-filter: blur(18px);
            border-top: 1px solid var(--line);
            padding: 0.45rem 0.65rem calc(0.55rem + env(safe-area-inset-bottom));
            z-index: 1000;
            box-shadow: 0 -14px 32px rgba(72, 41, 35, 0.08);
        }
        div[data-testid="stElementContainer"]:has(#mobile-bottom-nav-marker) + div[data-testid="stHorizontalBlock"] .stButton button,
        .element-container:has(#mobile-bottom-nav-marker) + div[data-testid="stHorizontalBlock"] .stButton button {
            min-height: 58px;
            border: 0;
            background: transparent;
            box-shadow: none;
            color: var(--muted);
            font-weight: 600;
            white-space: pre-line;
            line-height: 1.25;
            justify-content: center;
            border-radius: 16px;
        }
        div[data-testid="stElementContainer"]:has(#mobile-bottom-nav-marker) + div[data-testid="stHorizontalBlock"] .stButton button[kind="primary"],
        .element-container:has(#mobile-bottom-nav-marker) + div[data-testid="stHorizontalBlock"] .stButton button[kind="primary"] {
            color: var(--rose);
            background: transparent;
        }
        div[data-testid="stElementContainer"]:has(#mobile-bottom-nav-marker) + div[data-testid="stHorizontalBlock"] .stButton button[kind="primary"]::before,
        .element-container:has(#mobile-bottom-nav-marker) + div[data-testid="stHorizontalBlock"] .stButton button[kind="primary"]::before {
            content: "";
            position: absolute;
            top: 0;
            width: 58px;
            height: 4px;
            border-radius: 999px;
            background: var(--rose);
        }
    }
    @media (min-width: 721px) {
        .family-row-mobile:hover {
            transform: translateY(-1px);
            transition: transform .15s ease;
        }
        div[data-testid="stElementContainer"]:has(#mobile-bottom-nav-marker),
        .element-container:has(#mobile-bottom-nav-marker),
        div[data-testid="stElementContainer"]:has(#mobile-bottom-nav-marker) + div[data-testid="stHorizontalBlock"],
        .element-container:has(#mobile-bottom-nav-marker) + div[data-testid="stHorizontalBlock"] {
            display: none !important;
        }
    }
    </style>
    """,
    unsafe_allow_html=True,
)

if "guests" not in st.session_state:
    st.session_state.guests = load_data()["guests"]
if "page" not in st.session_state:
    st.session_state.page = "Convidados"
if "side_filter" not in st.session_state:
    st.session_state.side_filter = "Todos"
if "selected_family" not in st.session_state:
    st.session_state.selected_family = None

query_page = st.query_params.get("page")
if isinstance(query_page, list):
    query_page = query_page[0] if query_page else None
if query_page in ["Convidados", "Famílias", "Padrinhos"]:
    st.session_state.page = query_page

with st.sidebar:
    st.markdown(
        """
        <div style="display:flex; align-items:center; gap:.7rem; margin-bottom:2rem;">
            <div style="width:52px;height:52px;border:1px solid #eadfda;border-radius:50%;display:flex;align-items:center;justify-content:center;color:#c77d75;font-size:26px;">♡</div>
            <div style="font-family:'Playfair Display', Georgia, serif;color:#c77d75;font-size:1.45rem;font-weight:700;">Meu Casamento</div>
        </div>
        """,
        unsafe_allow_html=True,
    )
    st.button(
        "Convidados",
        use_container_width=True,
        type="primary" if st.session_state.page == "Convidados" else "secondary",
        on_click=set_page,
        args=("Convidados",),
    )
    st.button(
        "Famílias",
        use_container_width=True,
        type="primary" if st.session_state.page == "Famílias" else "secondary",
        on_click=set_page,
        args=("Famílias",),
    )
    st.button(
        "Padrinhos",
        use_container_width=True,
        type="primary" if st.session_state.page == "Padrinhos" else "secondary",
        on_click=set_page,
        args=("Padrinhos",),
    )
    st.markdown("<div style='height:38vh'></div>", unsafe_allow_html=True)
    st.caption("Tudo salvo automaticamente.")

total_guests = len([g for g in st.session_state.guests if g["name"].strip()])
favorite_count = len(
    [g for g in st.session_state.guests if g["name"].strip() and g["favorite"]]
)
families = family_summary()
family_count = len(families)

page_title = {
    "Convidados": "Convidados",
    "Famílias": "Famílias",
    "Padrinhos": "Padrinhos",
}[st.session_state.page]
page_subtitle = {
    "Convidados": "Organize sua lista do casamento",
    "Famílias": "Organize os convidados por núcleo familiar",
    "Padrinhos": "Lista especial do casamento",
}[st.session_state.page]

render_header(page_title, page_subtitle, {"Convidados": "♡", "Famílias": "👥", "Padrinhos": "♡"}[st.session_state.page])

if st.session_state.page == "Convidados":
    st.markdown('<div class="hero-card">', unsafe_allow_html=True)
    add_cols = st.columns([2.2, 1.35, 1, 1])
    new_name = add_cols[0].text_input(
        "Novo convidado",
        key="new_guest_name",
        placeholder="Digite o nome do convidado",
        label_visibility="collapsed",
    )
    family_options = ["Sem família"] + sorted(families.keys())
    new_family = add_cols[1].selectbox(
        "Família",
        family_options,
        key="new_guest_family",
        label_visibility="collapsed",
    )
    add_cols[2].selectbox(
        "Lado",
        ["Noiva", "Noivo"],
        key="new_guest_side",
        label_visibility="collapsed",
    )
    add_cols[3].button(
        "+ Adicionar",
        type="primary",
        use_container_width=True,
        on_click=add_guest_from_inputs,
    )
    st.markdown("</div>", unsafe_allow_html=True)

    count_cols = st.columns([1, 1, 1])
    count_cols[0].metric("Convidados", total_guests)
    count_cols[1].metric("Favoritos", favorite_count)
    count_cols[2].metric("Famílias", family_count)

    tool_cols = st.columns([1, 1])
    if tool_cols[0].button("Importar Pasta1.xlsx", use_container_width=True):
        if DEFAULT_XLSX.exists():
            import_xlsx(DEFAULT_XLSX)
            st.rerun()
        else:
            st.error("Nao encontrei a planilha.")
    tool_cols[1].download_button(
        "Baixar CSV",
        data=csv_bytes(),
        file_name="lista_casamento.csv",
        mime="text/csv",
        use_container_width=True,
    )

    filter_cols = st.columns([1, 1, 1, 2])
    for label, col in zip(["Todos", "Noiva", "Noivo"], filter_cols[:3]):
        if col.button(
            label,
            use_container_width=True,
            type="primary" if st.session_state.side_filter == label else "secondary",
        ):
            st.session_state.side_filter = label
            st.rerun()

    shown_guests = [
        guest
        for guest in st.session_state.guests
        if st.session_state.side_filter == "Todos"
        or guest.get("side", "Noiva") == st.session_state.side_filter
    ]

    st.markdown('<div class="table-card">', unsafe_allow_html=True)
    st.markdown(
        f'<div class="section-kicker">{len([g for g in shown_guests if g["name"].strip()])} convidados em {st.session_state.side_filter.lower()}</div>',
        unsafe_allow_html=True,
    )
    header = st.columns([2.5, 2, 1.15, 1, 0.45])
    header[0].markdown('<div class="table-header">Nome do convidado</div>', unsafe_allow_html=True)
    header[1].markdown('<div class="table-header">Família</div>', unsafe_allow_html=True)
    header[2].markdown('<div class="table-header">Status</div>', unsafe_allow_html=True)
    header[3].markdown('<div class="table-header">Lado</div>', unsafe_allow_html=True)
    header[4].markdown('<div class="table-header"></div>', unsafe_allow_html=True)
    for guest in shown_guests:
        guest_row(guest)
    st.markdown("</div>", unsafe_allow_html=True)

elif st.session_state.page == "Famílias":
    render_family_stats(family_count, total_guests)
    search_cols = st.columns([5, 0.7], vertical_alignment="center")
    search = search_cols[0].text_input(
        "Buscar família",
        placeholder="Buscar família",
        label_visibility="collapsed",
    ).strip().lower()
    search_cols[1].markdown('<div class="search-helper">☰</div>', unsafe_allow_html=True)
    if not families:
        st.info("Nenhum convidado cadastrado.")
    else:
        visible_families = []
        for family, values in sorted(families.items()):
            if search and search not in family.lower():
                continue
            visible_families.append((family, values))

        for family, values in visible_families:
            family_key = stable_key(family)
            render_family_row(family, values)
            action_cols = st.columns([1, 1], vertical_alignment="center")
            if action_cols[0].button("Editar família", key=f"edit_family_{family_key}", use_container_width=True):
                st.session_state.selected_family = None if st.session_state.selected_family == family else family
                st.session_state.selected_guest_id = None
                st.session_state.editing_guest_id = None
                st.rerun()
            if action_cols[1].button("Adicionar pessoa", key=f"quick_add_family_{family_key}", use_container_width=True):
                st.session_state.selected_family = family
                st.session_state.selected_guest_id = None
                st.session_state.editing_guest_id = None
                st.rerun()

            if st.session_state.selected_family == family:
                render_family_editor(family)

        if not visible_families:
            st.info("Nenhuma família encontrada com esse filtro.")

else:
    favorites = [
        guest for guest in st.session_state.guests if guest["name"].strip() and guest["favorite"]
    ]
    st.markdown(
        f"""
        <div class="soft-card">
            <div class="family-name">{favorite_count} padrinhos</div>
            <div class="family-sub">Pessoas especiais marcadas com estrela.</div>
        </div>
        """,
        unsafe_allow_html=True,
    )
    if st.button("+ Adicionar padrinho", type="primary", use_container_width=True):
        st.session_state.guests.insert(0, new_guest("", "Padrinhos", True))
        save_data()
        st.rerun()
    if not favorites:
        st.info("Nenhum favorito marcado ainda.")
    for guest in favorites:
        guest_row(guest, compact=True)

render_bottom_nav(st.session_state.page)
