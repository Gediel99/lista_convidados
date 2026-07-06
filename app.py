import csv
import io
import json
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
    first_name = names[0] if names else f"Familia {index}"
    return f"Familia {first_name}"


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
        "family": item.get("family", "Sem familia"),
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
    st.session_state.guests.insert(0, new_guest("", "Sem familia"))
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
        family = guest["family"].strip() or "Sem familia"
        families.setdefault(family, {"total": 0, "favorites": 0})
        families[family]["total"] += 1
        families[family]["favorites"] += int(guest["favorite"])
    return families


def guest_row(guest, compact=False):
    selected = st.session_state.get("selected_guest_id") == guest["id"]
    editing = st.session_state.get("editing_guest_id") == guest["id"]
    family_name = guest["family"].strip() or "Sem familia"

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
                "Familia",
                value=guest["family"],
                key=f"family_{guest['id']}",
                placeholder="Familia",
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


def set_page(page):
    st.session_state.page = page


st.set_page_config(page_title="Lista do casamento", layout="wide")

st.markdown(
    """
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Playfair+Display:wght@700&family=Inter:wght@400;500;600;700&display=swap');
    :root {
        --rose: #c77d75;
        --rose-dark: #b86d66;
        --ink: #171827;
        --muted: #858891;
        --line: #eadfda;
        --card: rgba(255, 255, 255, 0.88);
        --wash: #fbf6f2;
    }
    html, body, [data-testid="stAppViewContainer"] {
        background:
            radial-gradient(circle at 18% 4%, rgba(199, 125, 117, 0.11), transparent 30%),
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
    h1 {
        color: var(--rose);
        font-family: 'Playfair Display', Georgia, serif;
        font-size: clamp(2.45rem, 7vw, 4.35rem);
        line-height: 0.95;
        letter-spacing: 0;
        margin-bottom: 0.15rem;
    }
    h2, h3, p, label, div, span, input, button {
        font-family: 'Inter', system-ui, sans-serif;
    }
    [data-testid="stCaptionContainer"] {
        color: var(--muted);
        font-size: 1.02rem;
    }
    div[data-testid="stMetric"] {
        background: var(--card);
        border: 1px solid var(--line);
        border-radius: 8px;
        box-shadow: 0 18px 40px rgba(72, 41, 35, 0.08);
        padding: 14px 16px;
    }
    div[data-testid="stMetricValue"] {
        color: var(--ink);
        font-family: 'Playfair Display', Georgia, serif;
    }
    div[data-testid="stVerticalBlockBorderWrapper"] {
        border-color: var(--line);
        border-radius: 8px;
        background: var(--card);
        box-shadow: 0 10px 25px rgba(72, 41, 35, 0.06);
    }
    .stButton button {
        border-radius: 6px;
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
        border-radius: 8px;
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
        border-radius: 8px;
        background: rgba(255, 255, 255, 0.78);
        box-shadow: 0 18px 46px rgba(72, 41, 35, 0.08);
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
        border-radius: 8px !important;
    }
    .hero-card {
        border: 1px solid var(--line);
        background: var(--card);
        border-radius: 8px;
        padding: 1rem;
        box-shadow: 0 18px 46px rgba(72, 41, 35, 0.08);
        margin: 1.15rem 0 1.05rem;
    }
    .soft-card {
        border: 1px solid var(--line);
        background: linear-gradient(135deg, rgba(255,255,255,0.92), rgba(251,238,234,0.9));
        border-radius: 8px;
        padding: 1.1rem 1.2rem;
        box-shadow: 0 18px 46px rgba(72, 41, 35, 0.08);
        margin: 1rem 0;
    }
    .family-row {
        display: flex;
        align-items: center;
        justify-content: space-between;
        border: 1px solid var(--line);
        background: var(--card);
        border-radius: 8px;
        padding: 1rem;
        margin: 0.7rem 0;
        box-shadow: 0 10px 25px rgba(72, 41, 35, 0.05);
    }
    .family-name {
        color: var(--ink);
        font-family: 'Playfair Display', Georgia, serif;
        font-size: 1.35rem;
        font-weight: 700;
    }
    .family-sub {
        color: var(--muted);
        margin-top: 0.2rem;
    }
    .pill {
        background: linear-gradient(135deg, #dca8a0, #c8746f);
        color: white;
        border-radius: 8px;
        padding: 0.45rem 0.75rem;
        font-weight: 700;
        min-width: 2.5rem;
        text-align: center;
    }
    .bottom-nav {
        position: fixed;
        left: 50%;
        bottom: 0;
        transform: translateX(-50%);
        width: min(760px, 100vw);
        background: rgba(255,255,255,0.92);
        backdrop-filter: blur(18px);
        border-top: 1px solid var(--line);
        padding: 0.55rem 0.85rem 0.75rem;
        z-index: 99;
        box-shadow: 0 -14px 32px rgba(72, 41, 35, 0.08);
    }
    .bottom-nav-label {
        text-align: center;
        color: var(--muted);
        font-size: 0.86rem;
        margin-top: -0.35rem;
    }
    .section-kicker {
        color: var(--muted);
        font-weight: 600;
        margin: 0.35rem 0 0.8rem;
    }
    @media (max-width: 720px) {
        .block-container {
            padding: 1.4rem 0.9rem 5.8rem;
            max-width: 760px;
        }
        section[data-testid="stSidebar"] {
            display: none;
        }
        .hero-card {
            padding: 0.85rem;
            margin-top: 1rem;
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
            background: rgba(255, 255, 255, 0.9);
            border-radius: 8px;
            margin: 0.72rem 0;
            padding: 0.15rem;
        }
        div[data-testid="stVerticalBlockBorderWrapper"] div[data-testid="stHorizontalBlock"] {
            align-items: center;
            gap: 0.35rem;
        }
        div[data-testid="stVerticalBlockBorderWrapper"] .stButton button {
            min-height: 44px;
            border-radius: 8px;
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
            font-size: 0.78rem;
        }
        .family-row {
            padding: 0.85rem;
            margin: 0.65rem 0;
        }
    }
    @media (min-width: 900px) {
        .bottom-nav {
            display: none;
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

st.title(page_title)
st.caption(page_subtitle)

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
    st.markdown(
        f"""
        <div class="soft-card">
            <div style="display:flex; justify-content:space-around; gap:1rem; text-align:center;">
                <div><div class="family-name">{family_count}</div><div class="family-sub">famílias</div></div>
                <div style="width:1px; background:#e4d9d4;"></div>
                <div><div class="family-name">{total_guests}</div><div class="family-sub">convidados</div></div>
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )
    search = st.text_input(
        "Buscar família",
        placeholder="Buscar família",
        label_visibility="collapsed",
    ).strip().lower()
    if not families:
        st.info("Nenhum convidado cadastrado.")
    else:
        for family, values in sorted(families.items()):
            if search and search not in family.lower():
                continue
            st.markdown(
                f"""
                <div class="family-row">
                    <div>
                        <div class="family-name">{family}</div>
                        <div class="family-sub">{values['total']} convidado(s), {values['favorites']} favorito(s)</div>
                    </div>
                    <div class="pill">{values['total']}</div>
                </div>
                """,
                unsafe_allow_html=True,
            )

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

st.markdown('<div class="bottom-nav">', unsafe_allow_html=True)
nav_cols = st.columns(3)
nav_cols[0].button(
    "Convidados",
    use_container_width=True,
    type="primary" if st.session_state.page == "Convidados" else "secondary",
    on_click=set_page,
    args=("Convidados",),
)
nav_cols[1].button(
    "Famílias",
    use_container_width=True,
    type="primary" if st.session_state.page == "Famílias" else "secondary",
    on_click=set_page,
    args=("Famílias",),
)
nav_cols[2].button(
    "Padrinhos",
    use_container_width=True,
    type="primary" if st.session_state.page == "Padrinhos" else "secondary",
    on_click=set_page,
    args=("Padrinhos",),
)
st.markdown("</div>", unsafe_allow_html=True)
