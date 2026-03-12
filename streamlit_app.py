"""
Picking Lists — Blade B115  |  Application web Streamlit
"""

import io
import json
import subprocess
import tempfile
import traceback
from datetime import datetime
from pathlib import Path

import pandas as pd
import streamlit as st
import streamlit_authenticator as stauth
import yaml
from yaml.loader import SafeLoader

import auth_utils
import storage_utils as drive

# ── Config page ────────────────────────────────────────────────────────────────

st.set_page_config(
    page_title='Picking Lists — Blade B115',
    page_icon='📋',
    layout='wide',
    initial_sidebar_state='expanded',
)

VERSION = 'v3.7-web'
COMPONENTS = ['Blade', 'Blade service', 'PCW', 'Upper', 'Lower', 'WEB']

# ── CSS ────────────────────────────────────────────────────────────────────────

st.markdown("""
<style>
/* Reset Streamlit rounding */
div[data-testid="stButton"] button,
div[data-testid="stDownloadButton"] button,
div[data-testid="stForm"],
div[data-testid="stExpander"],
div[data-testid="stAlert"],
div[data-testid="stMetric"],
.stTextInput input, .stSelectbox select, .stNumberInput input {
    border-radius: 2px !important;
}

/* Compact buttons */
div[data-testid="stButton"] button {
    padding: 0.3rem 0.8rem;
    font-size: 0.85rem;
    font-weight: 600;
    letter-spacing: 0.02em;
}

/* Primary button — teal */
div[data-testid="stButton"] button[kind="primary"] {
    background: #00B4B4;
    border: none;
    color: #fff;
}
div[data-testid="stButton"] button[kind="primary"]:hover {
    background: #009898;
}

/* Header */
.sg-header {
    background: #2B2660;
    padding: 0.75rem 1.25rem;
    margin: -1rem -1rem 1.25rem -1rem;
    border-bottom: 3px solid #00B4B4;
}
.sg-header h2 { color: #fff; margin: 0 0 2px 0; font-size: 1.3rem; letter-spacing: -0.01em; }
.sg-header p  { color: #B8ADDE; margin: 0; font-size: 0.78rem; }
.sg-teal      { color: #00B4B4; }

/* Status badges */
.badge {
    display: inline-block;
    padding: 0.12rem 0.5rem;
    font-size: 0.72rem;
    font-weight: 700;
    letter-spacing: 0.04em;
    text-transform: uppercase;
}
.badge-ok   { background: #00B4B4; color: #fff; }
.badge-warn { background: #F5A623; color: #fff; }
.badge-err  { background: #D0021B; color: #fff; }

/* Log box */
.log-box {
    background: #14122a;
    color: #c8c0e8;
    font-family: 'JetBrains Mono', 'Fira Mono', monospace;
    font-size: 0.74rem;
    padding: 0.6rem 0.8rem;
    border-left: 3px solid #2B2660;
    max-height: 260px;
    overflow-y: auto;
    white-space: pre-wrap;
    line-height: 1.55;
}

/* Tabs — compact */
div[data-testid="stTabs"] button[data-baseweb="tab"] {
    font-weight: 600;
    font-size: 0.85rem;
    padding: 0.4rem 0.9rem;
}

/* Expanders */
details summary {
    font-weight: 600;
    font-size: 0.9rem;
}

/* Login card */
.login-card {
    background: #fff;
    border: 1px solid #ddd;
    border-top: 3px solid #2B2660;
    padding: 2rem;
    margin: 0 auto;
}

/* Admin panel */
.admin-row {
    background: #f8f7ff;
    border-left: 3px solid #00B4B4;
    padding: 0.4rem 0.7rem;
    margin-bottom: 0.4rem;
    font-size: 0.82rem;
}

/* Sidebar user name */
.sidebar-user {
    font-size: 0.95rem;
    font-weight: 700;
    color: #2B2660;
}
</style>
""", unsafe_allow_html=True)


# ── Auth ───────────────────────────────────────────────────────────────────────

def load_auth():
    with open('users.yaml') as f:
        return yaml.load(f, Loader=SafeLoader)


def _make_authenticator():
    cfg = load_auth()
    return stauth.Authenticate(
        cfg['credentials'],
        cfg['cookie']['name'],
        cfg['cookie']['key'],
        cfg['cookie']['expiry_days'],
    )


# Rechargé à chaque rerun pour prendre en compte les nouveaux utilisateurs approuvés
config        = load_auth()
authenticator = _make_authenticator()


# ── Login + Registration ────────────────────────────────────────────────────────

def _reset_session():
    """Efface toutes les clés de session liées à l'authentification."""
    for k in ['authentication_status', 'name', 'username',
              'main_xlsx_bytes', 'main_xlsx_loaded_at', 'log_lines']:
        st.session_state.pop(k, None)


def show_login():
    col1, col2, col3 = st.columns([1, 1.1, 1])
    with col2:
        st.markdown("""
        <div style='text-align:center;padding:1.5rem 0 1.2rem 0'>
            <div style='font-size:2.4rem'>📋</div>
            <h2 style='color:#2B2660;margin:0.3rem 0 0 0;letter-spacing:-0.02em'>Blade B115</h2>
            <p style='color:#888;font-size:0.82rem;margin-top:3px'>Picking Lists Generator</p>
        </div>
        """, unsafe_allow_html=True)

        login_tab, register_tab = st.tabs(['Connexion', 'Créer un compte'])

        with login_tab:
            try:
                authenticator.login(location='main')
            except Exception:
                _reset_session()
                st.rerun()

            status = st.session_state.get('authentication_status')
            if status is False:
                st.error('Identifiant ou mot de passe incorrect.')
            elif status is True:
                # Connecté via le formulaire → rerun pour passer à main_app
                st.rerun()

            # Bouton reset discret en cas de page blanche / cookie bloquant
            st.markdown('<br>', unsafe_allow_html=True)
            if st.button('Problème de connexion ? Réinitialiser la session',
                         key='reset_session',
                         help='Efface les cookies et la session locale'):
                _reset_session()
                st.rerun()

        with register_tab:
            _show_register_form()


def _show_register_form():
    with st.form('register_form', clear_on_submit=True):
        st.markdown('**Demander un accès**')
        username = st.text_input('Identifiant', placeholder='lettres et chiffres, min 3 caractères')
        name     = st.text_input('Nom complet', placeholder='ex: Jean Dupont')
        pwd      = st.text_input('Mot de passe', type='password', placeholder='min 6 caractères')
        pwd2     = st.text_input('Confirmer le mot de passe', type='password')
        submitted = st.form_submit_button('Envoyer la demande', use_container_width=True)

    if submitted:
        if pwd != pwd2:
            st.error('Les mots de passe ne correspondent pas.')
        else:
            ok, msg = auth_utils.register_user(username, name, pwd)
            if ok:
                st.success(msg)
            else:
                st.error(msg)


# ── Admin panel ────────────────────────────────────────────────────────────────

def _show_admin_panel():
    pending = auth_utils.load_pending()
    if not pending:
        st.caption('Aucune demande en attente.')
        return

    st.markdown(f'**{len(pending)} demande(s) en attente**')
    for p in pending:
        st.markdown(
            f'<div class="admin-row"><strong>{p["username"]}</strong> — {p["name"]}'
            f'<br><span style="color:#888;font-size:0.75rem">{p.get("requested_at","")}</span></div>',
            unsafe_allow_html=True,
        )
        c1, c2 = st.columns(2)
        with c1:
            if st.button('✅ Approuver', key=f'approve_{p["username"]}', use_container_width=True):
                if auth_utils.approve_user(p['username']):
                    st.success(f'{p["username"]} approuvé.')
                    st.rerun()
        with c2:
            if st.button('❌ Rejeter', key=f'reject_{p["username"]}', use_container_width=True):
                if auth_utils.reject_user(p['username']):
                    st.warning(f'{p["username"]} rejeté.')
                    st.rerun()


# ── Helpers data ───────────────────────────────────────────────────────────────

def get_main_xlsx() -> bytes | None:
    if 'main_xlsx_bytes' not in st.session_state:
        data = drive.download_file('main.xlsx')
        st.session_state['main_xlsx_bytes'] = data
        if data:
            st.session_state['main_xlsx_loaded_at'] = datetime.now().strftime('%d/%m/%Y à %H:%M')
    return st.session_state.get('main_xlsx_bytes')


def refresh_main_xlsx():
    for k in ['main_xlsx_bytes', 'main_xlsx_loaded_at']:
        st.session_state.pop(k, None)


def get_stock_cache() -> pd.DataFrame | None:
    data = drive.download_file('stock_cache.xlsx')
    if data is None:
        return None
    return pd.read_excel(
        io.BytesIO(data),
        dtype={'Handling Unit': str},
        parse_dates=['Shelf Life Expiration Date'],
    )


def log(msg: str):
    if 'log_lines' not in st.session_state:
        st.session_state['log_lines'] = []
    st.session_state['log_lines'].append(f"[{datetime.now().strftime('%H:%M:%S')}] {msg}")


# ── Helpers UI ─────────────────────────────────────────────────────────────────

def _libreoffice_available() -> bool:
    try:
        subprocess.run(['libreoffice', '--version'], capture_output=True, timeout=5)
        return True
    except Exception:
        return False


def _file_to_pdf(file_path: Path) -> Path | None:
    """Convertit un .xlsx ou .pptx en PDF via LibreOffice."""
    try:
        subprocess.run(
            ['libreoffice', '--headless', '--convert-to', 'pdf',
             str(file_path), '--outdir', str(file_path.parent)],
            capture_output=True, timeout=60
        )
        pdf = file_path.with_suffix('.pdf')
        return pdf if pdf.exists() else None
    except Exception:
        return None


def _download_buttons(files: dict, key_prefix: str):
    """Grille de boutons de téléchargement pour {fname: bytes}."""
    if not files:
        return
    cols = st.columns(min(len(files), 3))
    mime_map = {
        '.xlsx': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        '.pptx': 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
        '.pdf':  'application/pdf',
    }
    for i, (fname, fbytes) in enumerate(files.items()):
        with cols[i % 3]:
            mime = mime_map.get(Path(fname).suffix.lower(), 'application/octet-stream')
            st.download_button(
                f'⬇ {fname}', data=fbytes, file_name=fname, mime=mime,
                key=f'{key_prefix}_{fname}', use_container_width=True,
            )


def _component_selector(key_prefix: str):
    """Sélecteur de composants + PO. Retourne (selected, po_numbers)."""
    st.markdown('**Composants à générer**')
    cols = st.columns(3)
    selected = []
    for i, comp in enumerate(COMPONENTS):
        with cols[i % 3]:
            if st.checkbox(comp, value=True, key=f'{key_prefix}_comp_{comp}'):
                selected.append(comp)
    with st.expander('Numéros de PO (optionnel)'):
        po_cols = st.columns(3)
        po_numbers = {}
        for i, comp in enumerate(COMPONENTS):
            with po_cols[i % 3]:
                val = st.text_input(comp, key=f'{key_prefix}_po_{comp}',
                                    placeholder='ex: 4500123456')
                if val:
                    po_numbers[comp] = val
    return selected, po_numbers


class _Capture:
    """Capture stdout vers une liste de lignes."""
    def __init__(self):
        self.lines = []
    def write(self, t):
        if t.strip():
            self.lines.append(t.rstrip())
    def flush(self):
        pass


# ── Import SAP ─────────────────────────────────────────────────────────────────

def step_import_sap():
    loaded_at = st.session_state.get('main_xlsx_loaded_at', '')
    hint = f'main.xlsx chargé le {loaded_at}' if loaded_at else '⚠️ main.xlsx non chargé'

    with st.expander(f'📥 Import Stock SAP  ·  {hint}', expanded=True):
        uploaded = st.file_uploader(
            'Sélectionner l\'export SAP (.xlsx / .xls)',
            type=['xlsx', 'xls'], key='sap_upload',
        )
        if uploaded:
            if st.button('▶ Lancer l\'import', key='btn_import', type='primary'):
                _do_import_sap(uploaded)


def _do_import_sap(uploaded_file):
    import picking_list_generator as plg_mod
    import sys

    with tempfile.TemporaryDirectory() as tmpdir:
        tmpdir = Path(tmpdir)
        main_bytes = drive.download_file('main.xlsx')
        if main_bytes is None:
            st.error('main.xlsx introuvable sur le serveur.')
            return
        main_path = tmpdir / 'main.xlsx'
        main_path.write_bytes(main_bytes)
        sap_path = tmpdir / uploaded_file.name
        sap_path.write_bytes(uploaded_file.read())

        cap = _Capture()
        with st.status('Import en cours…', expanded=True) as status:
            old = sys.stdout
            sys.stdout = cap
            try:
                gen = plg_mod.PickingListGenerator(str(main_path))
                gen.import_stock_from_sap(str(sap_path))
            except Exception as e:
                sys.stdout = old
                status.update(label='❌ Erreur', state='error')
                st.error(str(e))
                log(f'ERREUR import SAP : {e}')
                return
            finally:
                sys.stdout = old
            status.update(label='✅ Import terminé', state='complete')

        drive.upload_file(main_path.read_bytes(), 'main.xlsx')
        drive.upload_file(json.dumps({
            'date': datetime.now().strftime('%d/%m/%Y à %H:%M'),
            'user': st.session_state.get('name', '?'),
        }, indent=2).encode(), 'last_import.json')
        cache_path = tmpdir / 'stock_cache.xlsx'
        if cache_path.exists():
            drive.upload_file(cache_path.read_bytes(), 'stock_cache.xlsx')

        for line in cap.lines:
            log(line)
        log(f'✅ Import terminé par {st.session_state.get("name", "?")}')
        refresh_main_xlsx()
        st.rerun()


# ── Générer Picking Lists ──────────────────────────────────────────────────────

def step_generate_picking():
    has_pdf = _libreoffice_available()
    with st.expander('📋 Générer les Picking Lists', expanded=False):
        selected, po_numbers = _component_selector('picking')
        print_pdf = has_pdf and st.checkbox(
            '🖨 Générer aussi en PDF (impression directe)', key='picking_pdf')
        c1, c2 = st.columns([3, 1])
        with c2:
            if st.button('▶ Générer', key='btn_picking', type='primary', use_container_width=True):
                if not selected:
                    st.warning('Sélectionnez au moins un composant.')
                else:
                    _do_generate_picking(selected, po_numbers, print_pdf)
        if 'generated_picking_files' in st.session_state:
            st.divider()
            st.markdown('**Fichiers générés**')
            _download_buttons(st.session_state['generated_picking_files'], 'dl_pl')


def _do_generate_picking(selected_comps, po_numbers, print_pdf=False, *, _rerun=True):
    import picking_list_generator as plg_mod
    import sys

    main_bytes = get_main_xlsx()
    if main_bytes is None:
        st.error('main.xlsx introuvable.')
        return False

    with tempfile.TemporaryDirectory() as tmpdir:
        tmpdir = Path(tmpdir)
        main_path = tmpdir / 'main.xlsx'
        main_path.write_bytes(main_bytes)
        counter_bytes = drive.download_file('pl_counter.json')
        if counter_bytes:
            (tmpdir / 'pl_counter.json').write_bytes(counter_bytes)
        picking_out = tmpdir / 'picking_lists'
        picking_out.mkdir()

        cap = _Capture()
        with st.status('Génération des picking lists…', expanded=True) as status:
            old = sys.stdout
            sys.stdout = cap
            try:
                gen = plg_mod.PickingListGenerator(str(main_path))
                gen.load_data()
                gen.remove_staging_locations(target_file=str(main_path))
                gen.generate_picking_lists(components_filter=selected_comps)
                gen.save_picking_lists(output_folder=str(picking_out), shared_dir=tmpdir)
                gen.save_updated_stock(output_file=str(main_path))
            except Exception as e:
                sys.stdout = old
                status.update(label='❌ Erreur', state='error')
                st.error(str(e))
                log(f'ERREUR picking : {e}\n{traceback.format_exc()}')
                return False
            finally:
                sys.stdout = old
            status.update(label='✅ Picking lists générées', state='complete')

        pl_files = sorted(picking_out.glob('PL_*.xlsx'))
        if not pl_files:
            st.warning('Aucune picking list générée.')
            return False

        drive.upload_file(main_path.read_bytes(), 'main.xlsx')
        counter_path = tmpdir / 'pl_counter.json'
        if counter_path.exists():
            drive.upload_file(counter_path.read_bytes(), 'pl_counter.json')
        cache_path = tmpdir / 'stock_cache.xlsx'
        if cache_path.exists():
            drive.upload_file(cache_path.read_bytes(), 'stock_cache.xlsx')

        pl_folder_id = drive.get_subfolder_id('picking_lists')
        generated = {}
        for f in pl_files:
            fbytes = f.read_bytes()
            drive.upload_file(fbytes, f.name, pl_folder_id)
            generated[f.name] = fbytes
            if print_pdf:
                pdf = _file_to_pdf(f)
                if pdf:
                    generated[pdf.name] = pdf.read_bytes()

        if po_numbers:
            drive.upload_file(
                json.dumps(po_numbers, indent=2, ensure_ascii=False).encode(),
                'po_numbers.json')

        st.session_state['generated_picking_files'] = generated
        for line in cap.lines:
            log(line)
        log(f'✅ {len(pl_files)} picking list(s) générée(s)')
        refresh_main_xlsx()
        if _rerun:
            st.rerun()
        return True


# ── Mettre à jour les PowerPoints ─────────────────────────────────────────────

def step_update_pptx():
    has_pdf = _libreoffice_available()
    with st.expander('📊 Mettre à jour les PowerPoints', expanded=False):
        st.caption('Les fichiers sources .pptx doivent être dans /data/sources/')
        print_pdf = has_pdf and st.checkbox(
            '🖨 Générer aussi en PDF (impression directe)', key='pptx_pdf')
        c1, c2 = st.columns([3, 1])
        with c2:
            if st.button('▶ Lancer', key='btn_pptx', type='primary', use_container_width=True):
                _do_update_pptx(print_pdf)
        if 'generated_pptx_files' in st.session_state:
            st.divider()
            st.markdown('**Fichiers générés**')
            _download_buttons(st.session_state['generated_pptx_files'], 'dl_pptx')


def _do_update_pptx(print_pdf=False, *, _rerun=True):
    import update_all_powerpoints
    import sys

    main_bytes = get_main_xlsx()
    if main_bytes is None:
        st.error('main.xlsx introuvable.')
        return False

    with tempfile.TemporaryDirectory() as tmpdir:
        tmpdir = Path(tmpdir)
        main_path = tmpdir / 'main.xlsx'
        main_path.write_bytes(main_bytes)

        sources_folder_id = drive.get_subfolder_id('sources')
        source_files = drive.list_files(sources_folder_id, pattern='.pptx')
        if not source_files:
            st.error('Aucun fichier .pptx dans /data/sources/.')
            return False

        sources_dir = tmpdir / 'sources'
        sources_dir.mkdir()
        for f in source_files:
            fbytes = drive.download_file(f['name'], sources_folder_id)
            if fbytes:
                (sources_dir / f['name']).write_bytes(fbytes)

        po_bytes = drive.download_file('po_numbers.json')
        po_numbers = json.loads(po_bytes.decode()) if po_bytes else {}

        (tmpdir / 'picking_lists').mkdir(exist_ok=True)
        pl_folder_id = drive.get_subfolder_id('picking_lists')
        for f in drive.list_files(pl_folder_id, pattern='PL_'):
            fbytes = drive.download_file(f['name'], pl_folder_id)
            if fbytes:
                (tmpdir / 'picking_lists' / f['name']).write_bytes(fbytes)

        pptx_out = tmpdir / 'powerpoints_updated'
        pptx_out.mkdir()

        cap = _Capture()
        with st.status('Mise à jour des PowerPoints…', expanded=True) as status:
            old = sys.stdout
            sys.stdout = cap
            try:
                update_all_powerpoints.main(po_numbers=po_numbers or None, shared_dir=tmpdir)
            except Exception as e:
                sys.stdout = old
                status.update(label='❌ Erreur', state='error')
                st.error(str(e))
                log(f'ERREUR pptx : {e}\n{traceback.format_exc()}')
                return False
            finally:
                sys.stdout = old
            status.update(label='✅ PowerPoints mis à jour', state='complete')

        pw_files = sorted(pptx_out.glob('PW_*.pptx'))
        if not pw_files:
            st.warning('Aucun PowerPoint généré.')
            return False

        pptx_folder_id = drive.get_subfolder_id('powerpoints_updated')
        generated = {}
        for f in pw_files:
            fbytes = f.read_bytes()
            drive.upload_file(fbytes, f.name, pptx_folder_id)
            generated[f.name] = fbytes
            if print_pdf:
                pdf = _file_to_pdf(f)
                if pdf:
                    generated[pdf.name] = pdf.read_bytes()

        drive.upload_file(b'{}', 'po_numbers.json')
        st.session_state['generated_pptx_files'] = generated
        for line in cap.lines:
            log(line)
        log(f'✅ {len(pw_files)} PowerPoint(s) mis à jour')
        if _rerun:
            st.rerun()
        return True


# ── Archiver ───────────────────────────────────────────────────────────────────

def step_archive():
    with st.expander('📦 Archiver les Picking Lists', expanded=False):
        st.caption('Déplace les PL_*.xlsx du dossier picking_lists vers une archive numérotée')
        if st.button('▶ Archiver', key='btn_archive', type='primary'):
            _do_archive()


def _do_archive():
    import shutil

    pl_folder_id = drive.get_subfolder_id('picking_lists')
    pl_files = drive.list_files(pl_folder_id, pattern='PL_')
    pl_files = [f for f in pl_files if f['name'].startswith('PL_') and f['name'].endswith('.xlsx')]

    if not pl_files:
        st.info('Aucune picking list à archiver.')
        return

    archive_index_bytes = drive.download_file('archive_index.json')
    archive_index = (json.loads(archive_index_bytes.decode())
                     if archive_index_bytes else {'archives': [], 'next_number': 1})
    archive_num  = archive_index.get('next_number', 1)
    archive_name = f'#{archive_num:04d}_{datetime.now().strftime("%Y-%m-%d")}'

    archive_folder_id = drive.get_subfolder_id(
        archive_name, drive.get_subfolder_id('picking_lists_archive'))
    archived_names = []
    for f in pl_files:
        shutil.move(f['id'], str(Path(archive_folder_id) / f['name']))
        archived_names.append(f['name'])

    archive_index['archives'].insert(0, {
        'number':    archive_num,
        'folder':    archive_name,
        'date':      datetime.now().strftime('%d/%m/%Y'),
        'datetime':  datetime.now().strftime('%d/%m/%Y à %H:%M'),
        'user':      st.session_state.get('name', '?'),
        'files':     archived_names,
        'components': [],
    })
    archive_index['next_number'] = archive_num + 1
    drive.upload_file(json.dumps(archive_index, indent=2).encode(), 'archive_index.json')

    st.session_state.pop('generated_picking_files', None)
    log(f'✅ {len(pl_files)} fichier(s) archivé(s) dans {archive_name}.')
    st.success(f'{len(pl_files)} fichier(s) archivé(s).')
    st.rerun()


# ── Picking Ad-Hoc ─────────────────────────────────────────────────────────────

def tab_adhoc():
    col_left, col_right = st.columns([1, 1], gap='large')

    with col_left:
        st.markdown('#### 🔧 Picking Ad-Hoc')
        st.caption('Génère une picking list à la demande pour des références précises')

        main_bytes = get_main_xlsx()
        all_refs, ref_desc_map = [], {}
        if main_bytes:
            try:
                for sheet in COMPONENTS:
                    df = pd.read_excel(io.BytesIO(main_bytes), sheet_name=sheet, header=0)
                    for _, row in df.iterrows():
                        ref = str(row.iloc[0]).strip()
                        if ref and ref != 'nan' and ref not in all_refs:
                            desc = str(row.iloc[2]).strip() if len(row) > 2 and pd.notna(row.iloc[2]) else ''
                            if desc and desc != 'nan':
                                ref_desc_map[ref] = desc
                            all_refs.append(ref)
            except Exception:
                pass

        if 'adhoc_items' not in st.session_state:
            st.session_state['adhoc_items'] = []

        with st.form('adhoc_form', clear_on_submit=True):
            ref_options = [f"{r}  —  {ref_desc_map[r]}" if r in ref_desc_map else r for r in all_refs]
            ref_input = st.selectbox('Référence', options=[''] + ref_options,
                                     index=0, label_visibility='collapsed')
            c1, c2, c3 = st.columns([2, 1, 1])
            with c1:
                qty_input = st.number_input('Quantité', min_value=0.1, value=1.0, step=1.0,
                                            format='%g', label_visibility='collapsed')
            with c2:
                unit_input = st.radio('Unité', ['palette', 'pièce'],
                                      label_visibility='collapsed')
            with c3:
                add_btn = st.form_submit_button('+ Ajouter', use_container_width=True)
            if add_btn and ref_input:
                pure = ref_input.split('  —  ')[0].strip().upper()
                st.session_state['adhoc_items'].append(
                    {'reference': pure, 'quantity': qty_input, 'unit': unit_input})

        items = st.session_state['adhoc_items']
        if items:
            st.dataframe(pd.DataFrame(items), hide_index=True, use_container_width=True)
            c1, c2 = st.columns(2)
            with c1:
                if st.button('🗑 Vider la liste', key='adhoc_clear', use_container_width=True):
                    st.session_state['adhoc_items'] = []
                    st.rerun()
            with c2:
                if st.button('▶ Générer', key='adhoc_gen', type='primary', use_container_width=True):
                    _do_adhoc_picking(list(items))
        else:
            st.caption('_Aucun article — ajoutez une référence ci-dessus._')

        if 'adhoc_result' in st.session_state:
            fname, fbytes = st.session_state['adhoc_result']
            st.success('Picking list générée.')
            st.download_button(
                f'⬇ Télécharger {fname}', data=fbytes, file_name=fname,
                mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                key='dl_adhoc', use_container_width=True,
            )

    with col_right:
        st.markdown('#### 🔍 Consulter le stock')
        st.caption('Rechercher une référence et afficher ses emplacements et quantités')

        cache_df = get_stock_cache()
        if cache_df is None:
            st.warning('Aucun cache. Lancez d\'abord un import SAP.')
        else:
            st.caption(f'{len(cache_df)} lignes dans le cache')
            search = st.text_input('Référence', placeholder='ex: 1234567-001',
                                   key='stock_search', label_visibility='collapsed')
            if search:
                mask = cache_df['Product'].astype(str).str.upper().str.contains(search.upper())
                result = cache_df[mask]
                if result.empty:
                    st.info('Aucun résultat.')
                else:
                    cols_show = ['Handling Unit', 'Storage Bin', 'Quantity',
                                 'Base Unit of Measure', 'Shelf Life Expiration Date']
                    cols_show = [c for c in cols_show if c in result.columns]
                    st.dataframe(result[cols_show], use_container_width=True, hide_index=True)
                    if 'Quantity' in result.columns:
                        st.metric('Quantité totale', f'{result["Quantity"].sum():g}')


def _do_adhoc_picking(items):
    import picking_list_generator as plg_mod
    import sys

    main_bytes = get_main_xlsx()
    if main_bytes is None:
        st.error('main.xlsx introuvable.')
        return

    with tempfile.TemporaryDirectory() as tmpdir:
        tmpdir = Path(tmpdir)
        main_path = tmpdir / 'main.xlsx'
        main_path.write_bytes(main_bytes)
        counter_bytes = drive.download_file('pl_counter.json')
        if counter_bytes:
            (tmpdir / 'pl_counter.json').write_bytes(counter_bytes)
        picking_out = tmpdir / 'picking_lists'
        picking_out.mkdir()

        cap = _Capture()
        old = sys.stdout
        sys.stdout = cap
        try:
            with st.spinner('Génération…'):
                gen = plg_mod.PickingListGenerator(str(main_path))
                gen.load_data()
                filename = gen.generate_adhoc_picking_list(
                    items, output_folder=str(picking_out), shared_dir=tmpdir)
                gen.save_updated_stock(output_file=str(main_path))
        except Exception as e:
            sys.stdout = old
            st.error(str(e))
            return
        finally:
            sys.stdout = old

        for line in cap.lines:
            log(line)

        if filename and Path(filename).exists():
            fbytes = Path(filename).read_bytes()
            fname  = Path(filename).name
            pl_folder_id = drive.get_subfolder_id('picking_lists')
            drive.upload_file(fbytes, fname, pl_folder_id)
            drive.upload_file(main_path.read_bytes(), 'main.xlsx')
            counter_path = tmpdir / 'pl_counter.json'
            if counter_path.exists():
                drive.upload_file(counter_path.read_bytes(), 'pl_counter.json')
            cache_path = tmpdir / 'stock_cache.xlsx'
            if cache_path.exists():
                drive.upload_file(cache_path.read_bytes(), 'stock_cache.xlsx')
            st.session_state['adhoc_result'] = (fname, fbytes)
            st.session_state['adhoc_items'] = []
            refresh_main_xlsx()
            log(f'✅ Ad-hoc généré : {fname}')
            st.rerun()
        else:
            st.warning('Aucun fichier généré.')


# ── BOM ────────────────────────────────────────────────────────────────────────

def tab_bom():
    col_left, col_right = st.columns([1, 1], gap='large')

    with col_left:
        st.markdown('#### 📑 Générer l\'onglet BOM')
        st.caption('Consolider toutes les références dans un onglet dédié de main.xlsx')
        if st.button('▶ Générer l\'onglet BOM', key='btn_bom_sheet', type='primary'):
            _do_bom_sheet()

    with col_right:
        st.markdown('#### 🔄 Synchroniser le BOM')
        st.caption('Propager les changements de main.xlsx vers les PowerPoint sources')
        c1, c2 = st.columns(2)
        with c1:
            if st.button('🔍 Simuler', key='btn_sim_bom', use_container_width=True):
                _do_sync_bom(dry_run=True)
        with c2:
            if st.button('▶ Synchroniser', key='btn_sync_bom',
                         type='primary', use_container_width=True):
                _do_sync_bom(dry_run=False)
        if 'bom_report' in st.session_state:
            fname, fbytes = st.session_state['bom_report']
            st.download_button(f'⬇ Rapport {fname}', data=fbytes,
                               file_name=fname, mime='text/html', key='dl_bom')


def _do_bom_sheet():
    import picking_list_generator as plg_mod
    import sys

    main_bytes = get_main_xlsx()
    if main_bytes is None:
        st.error('main.xlsx introuvable.')
        return
    with tempfile.TemporaryDirectory() as tmpdir:
        tmpdir = Path(tmpdir)
        main_path = tmpdir / 'main.xlsx'
        main_path.write_bytes(main_bytes)
        cap = _Capture()
        old = sys.stdout
        sys.stdout = cap
        try:
            with st.spinner('Génération onglet BOM…'):
                plg_mod.PickingListGenerator(str(main_path)).generate_bom_sheet()
        except Exception as e:
            sys.stdout = old
            st.error(str(e))
            return
        finally:
            sys.stdout = old
        for line in cap.lines:
            log(line)
        drive.upload_file(main_path.read_bytes(), 'main.xlsx')
        refresh_main_xlsx()
        log('✅ Onglet BOM généré')
        st.success('Onglet BOM généré.')
        st.rerun()


def _do_sync_bom(dry_run=False):
    import sync_bom
    import sys

    main_bytes = get_main_xlsx()
    if main_bytes is None:
        st.error('main.xlsx introuvable.')
        return
    sources_folder_id = drive.get_subfolder_id('sources')

    with tempfile.TemporaryDirectory() as tmpdir:
        tmpdir = Path(tmpdir)
        main_path = tmpdir / 'main.xlsx'
        main_path.write_bytes(main_bytes)
        sources_dir = tmpdir / 'sources'
        sources_dir.mkdir()
        for f in drive.list_files(sources_folder_id):
            fbytes = drive.download_file(f['name'], sources_folder_id)
            if fbytes:
                (sources_dir / f['name']).write_bytes(fbytes)

        old_excel   = sync_bom.EXCEL_FILE
        old_sources = sync_bom.SOURCES_DIR
        sync_bom.EXCEL_FILE  = main_path
        sync_bom.SOURCES_DIR = sources_dir

        cap = _Capture()
        old = sys.stdout
        sys.stdout = cap
        try:
            label = 'Simulation' if dry_run else 'Synchronisation'
            with st.spinner(f'{label} BOM…'):
                report_path = sync_bom.main(dry_run=dry_run)
        except Exception as e:
            sys.stdout = old
            sync_bom.EXCEL_FILE  = old_excel
            sync_bom.SOURCES_DIR = old_sources
            st.error(str(e))
            return
        finally:
            sys.stdout = old
            sync_bom.EXCEL_FILE  = old_excel
            sync_bom.SOURCES_DIR = old_sources

        for line in cap.lines:
            log(line)
        if not dry_run:
            for f in sources_dir.glob('*.pptx'):
                drive.upload_file(f.read_bytes(), f.name, sources_folder_id)
        if report_path and Path(report_path).exists():
            rname  = Path(report_path).name
            rbytes = Path(report_path).read_bytes()
            drive.upload_file(rbytes, rname)
            st.session_state['bom_report'] = (rname, rbytes)
        action = 'Simulation terminée' if dry_run else 'Synchronisation terminée'
        st.success(f'{action}.')
        st.rerun()


# ── Workflow complet ───────────────────────────────────────────────────────────

def tab_workflow():
    with st.container():
        st.markdown('#### ▶ Workflow complet')
        st.caption('Enchaîne Import SAP → Picking Lists → PowerPoints en une seule action')

        if st.button('▶  Lancer le workflow complet', key='btn_wf_all',
                     type='primary', use_container_width=True):
            st.session_state['show_wf_config'] = True

        if st.session_state.get('show_wf_config'):
            with st.container():
                st.divider()
                sap_file = st.file_uploader(
                    'Export SAP', type=['xlsx', 'xls'], key='wf_sap')
                selected, po_numbers = _component_selector('wf')
                c1, c2 = st.columns([1, 3])
                with c1:
                    if st.button('Annuler', key='wf_cancel'):
                        st.session_state['show_wf_config'] = False
                        st.rerun()
                with c2:
                    if st.button('▶ Lancer', key='wf_confirm', type='primary',
                                 use_container_width=True):
                        if not sap_file:
                            st.warning('Sélectionnez un fichier SAP.')
                        elif not selected:
                            st.warning('Sélectionnez au moins un composant.')
                        else:
                            st.session_state['show_wf_config'] = False
                            _do_import_sap(sap_file)

    st.divider()

    st.markdown('#### Étapes individuelles')
    step_import_sap()
    step_generate_picking()
    step_update_pptx()
    step_archive()


# ── Page principale ────────────────────────────────────────────────────────────

def main_app():
    user     = st.session_state.get('name', '')
    username = st.session_state.get('username', '')

    # ── Sidebar ───────────────────────────────────────────────────────────────
    with st.sidebar:
        st.markdown(f'<div class="sidebar-user">👤 {user}</div>', unsafe_allow_html=True)
        authenticator.logout('Déconnexion', 'sidebar', key='logout')
        st.divider()

        loaded_at = st.session_state.get('main_xlsx_loaded_at', '')
        if loaded_at:
            st.markdown(f'<span class="badge badge-ok">main.xlsx</span> chargé {loaded_at}',
                        unsafe_allow_html=True)
        else:
            st.markdown('<span class="badge badge-warn">main.xlsx</span> non chargé',
                        unsafe_allow_html=True)

        has_cache = drive.download_file('stock_cache.xlsx') is not None
        if has_cache:
            st.markdown('<span class="badge badge-ok">Cache stock</span> disponible',
                        unsafe_allow_html=True)
        else:
            st.markdown('<span class="badge badge-warn">Cache stock</span> absent',
                        unsafe_allow_html=True)

        st.divider()

        # Admin panel — visible uniquement pour l'admin
        if auth_utils.is_admin(username):
            with st.expander('🔑 Gestion des accès', expanded=False):
                _show_admin_panel()
            st.divider()

        st.caption(VERSION)
        st.divider()

        st.markdown('**📋 Journal**')
        lines = st.session_state.get('log_lines', [])
        if lines:
            content = '\n'.join(lines[-60:])
            st.markdown(f'<div class="log-box">{content}</div>', unsafe_allow_html=True)
            if st.button('🗑 Effacer', key='clear_log'):
                st.session_state['log_lines'] = []
                st.rerun()
        else:
            st.caption('_Aucune activité_')

    # ── Header ────────────────────────────────────────────────────────────────
    st.markdown(f"""
    <div class="sg-header">
        <h2>📋 Gestion des Picking Lists &nbsp;
            <span class="sg-teal">BLADE B115</span>
        </h2>
        <p>{VERSION} &nbsp;·&nbsp; {user}</p>
    </div>
    """, unsafe_allow_html=True)

    # ── Tabs ──────────────────────────────────────────────────────────────────
    tab1, tab2, tab3 = st.tabs([
        '📋  Workflow de production',
        '🔧  Consultation & Ad-Hoc',
        '📊  Gestion du BOM',
    ])

    with tab1:
        tab_workflow()

    with tab2:
        tab_adhoc()

    with tab3:
        tab_bom()


# ── Point d'entrée ─────────────────────────────────────────────────────────────

_status = st.session_state.get('authentication_status')
_name   = st.session_state.get('name')

if _status is True and _name:
    main_app()
elif _status is True and not _name:
    # Cookie présent mais session incohérente (ex: redémarrage serveur)
    _reset_session()
    st.rerun()
else:
    show_login()
