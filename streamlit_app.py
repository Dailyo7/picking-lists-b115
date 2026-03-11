"""
Picking Lists — Blade B115  |  Application web Streamlit
"""

import io
import json
import tempfile
import traceback
from datetime import datetime
from pathlib import Path

import pandas as pd
import streamlit as st
import streamlit_authenticator as stauth
import yaml
from yaml.loader import SafeLoader

import drive_utils as drive

# ── Config page ───────────────────────────────────────────────────────────────

st.set_page_config(
    page_title='Picking Lists — Blade B115',
    page_icon='📋',
    layout='wide',
    initial_sidebar_state='collapsed',
)

VERSION = 'v3.6-web'
COMPONENTS = ['Blade', 'Blade service', 'PCW', 'Upper', 'Lower', 'WEB']

# ── CSS Siemens Gamesa ─────────────────────────────────────────────────────────

st.markdown("""
<style>
/* Header */
.sg-header {
    background: #2B2660;
    padding: 1.2rem 2rem;
    margin: -1rem -1rem 1.5rem -1rem;
    border-bottom: 3px solid #00B4B4;
}
.sg-header h1 { color: white; margin: 0; font-size: 1.5rem; }
.sg-header p  { color: #B8ADDE; margin: 0; font-size: 0.85rem; }
.sg-teal      { color: #00B4B4; font-weight: bold; }

/* Sections */
.sg-section {
    display: flex;
    align-items: center;
    gap: 0.5rem;
    margin: 1.5rem 0 0.75rem 0;
    border-bottom: 1px solid #D0D0DE;
    padding-bottom: 0.4rem;
}
.sg-section-bar { width: 3px; height: 16px; background: #00B4B4; border-radius: 2px; }
.sg-section-lbl { color: #2B2660; font-weight: 700; font-size: 0.8rem; letter-spacing: 0.08em; text-transform: uppercase; }

/* Cards */
.sg-card {
    background: white;
    border: 1px solid #D0D0DE;
    border-radius: 8px;
    padding: 1rem 1.2rem;
    margin-bottom: 0.6rem;
}

/* Status bar */
.sg-statusbar {
    background: #2B2660;
    color: #B8ADDE;
    padding: 0.4rem 1rem;
    font-size: 0.75rem;
    position: fixed;
    bottom: 0; left: 0; right: 0;
    z-index: 999;
}

/* Log */
.log-box {
    background: #1A1840;
    color: #E8E0F8;
    font-family: monospace;
    font-size: 0.8rem;
    padding: 0.8rem;
    border-radius: 6px;
    max-height: 200px;
    overflow-y: auto;
    white-space: pre-wrap;
}

button[kind="primary"] { background: #00B4B4 !important; color: black !important; }
</style>
""", unsafe_allow_html=True)


# ── Auth ──────────────────────────────────────────────────────────────────────

def load_auth():
    with open('users.yaml') as f:
        return yaml.load(f, Loader=SafeLoader)

config = load_auth()

authenticator = stauth.Authenticate(
    config['credentials'],
    config['cookie']['name'],
    config['cookie']['key'],
    config['cookie']['expiry_days'],
)


# ── Login ─────────────────────────────────────────────────────────────────────

def show_login():
    col1, col2, col3 = st.columns([1, 1.2, 1])
    with col2:
        st.markdown("""
        <div style='text-align:center; padding: 2rem 0 1rem 0;'>
            <div style='font-size:2.5rem;'>📋</div>
            <h2 style='color:#2B2660; margin:0;'>Blade B115</h2>
            <p style='color:#666; font-size:0.9rem;'>Picking Lists Generator</p>
        </div>
        """, unsafe_allow_html=True)
        authenticator.login(location='main')
        if st.session_state.get('authentication_status') is False:
            st.error('Identifiant ou mot de passe incorrect.')
        elif st.session_state.get('authentication_status') is None:
            st.info('Veuillez vous connecter.')


# ── Session helpers ────────────────────────────────────────────────────────────

def get_main_xlsx() -> bytes | None:
    """Télécharge main.xlsx depuis Drive (cache de session)."""
    if 'main_xlsx_bytes' not in st.session_state:
        with st.spinner('Chargement de main.xlsx depuis Google Drive…'):
            data = drive.download_file('main.xlsx')
        st.session_state['main_xlsx_bytes'] = data
        if data:
            st.session_state['main_xlsx_loaded_at'] = datetime.now().strftime('%d/%m/%Y à %H:%M')
    return st.session_state.get('main_xlsx_bytes')


def refresh_main_xlsx():
    """Force le rechargement de main.xlsx depuis Drive."""
    for k in ['main_xlsx_bytes', 'main_xlsx_loaded_at']:
        st.session_state.pop(k, None)
    drive.get_folder_id.clear()


def get_stock_cache() -> pd.DataFrame | None:
    """Télécharge stock_cache.xlsx depuis Drive."""
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


def show_log():
    lines = st.session_state.get('log_lines', [])
    if lines:
        content = '\n'.join(lines[-50:])
        st.markdown(f'<div class="log-box">{content}</div>', unsafe_allow_html=True)
        if st.button('Effacer le journal', key='clear_log'):
            st.session_state['log_lines'] = []
            st.rerun()


# ── Sections UI ───────────────────────────────────────────────────────────────

def section(label: str):
    st.markdown(f"""
    <div class="sg-section">
        <div class="sg-section-bar"></div>
        <span class="sg-section-lbl">{label}</span>
    </div>
    """, unsafe_allow_html=True)


# ── Étape : Import Stock SAP ───────────────────────────────────────────────────

def step_import_sap():
    with st.container():
        st.markdown('<div class="sg-card">', unsafe_allow_html=True)
        c1, c2 = st.columns([4, 1])
        with c1:
            st.markdown('**📥 Import Stock SAP**')
            loaded_at = st.session_state.get('main_xlsx_loaded_at', '')
            hint = f'main.xlsx chargé le {loaded_at}' if loaded_at else 'main.xlsx non chargé'
            st.caption(f'Importer un export SAP dans main.xlsx  ·  {hint}')
        with c2:
            uploaded = st.file_uploader(
                'Export SAP', type=['xlsx', 'xls'],
                label_visibility='collapsed', key='sap_upload'
            )

        if uploaded:
            if st.button('▶ Lancer l\'import', key='btn_import', type='primary'):
                _do_import_sap(uploaded)
        st.markdown('</div>', unsafe_allow_html=True)


def _do_import_sap(uploaded_file):
    import picking_list_generator as plg_mod

    with tempfile.TemporaryDirectory() as tmpdir:
        tmpdir = Path(tmpdir)

        # Récupérer main.xlsx depuis Drive
        main_bytes = drive.download_file('main.xlsx')
        if main_bytes is None:
            st.error('main.xlsx introuvable sur Google Drive.')
            return

        main_path = tmpdir / 'main.xlsx'
        main_path.write_bytes(main_bytes)

        # Écrire le fichier SAP uploadé
        sap_path = tmpdir / uploaded_file.name
        sap_path.write_bytes(uploaded_file.read())

        log_lines = []
        try:
            with st.spinner('Import en cours…'):
                # Capturer le stdout
                import sys
                class _Capture:
                    def write(self, t): log_lines.append(t.rstrip())
                    def flush(self): pass

                old_out = sys.stdout
                sys.stdout = _Capture()
                try:
                    gen = plg_mod.PickingListGenerator(str(main_path))
                    gen.import_stock_from_sap(str(sap_path))
                finally:
                    sys.stdout = old_out

            # Sauvegarder main.xlsx mis à jour sur Drive
            drive.upload_file(main_path.read_bytes(), 'main.xlsx')

            # Sauvegarder last_import.json
            import_info = json.dumps({
                'date': datetime.now().strftime('%d/%m/%Y à %H:%M'),
                'user': st.session_state.get('name', '?'),
            }, indent=2)
            drive.upload_file(import_info.encode(), 'last_import.json')

            # Sauvegarder stock_cache.xlsx si généré
            cache_path = tmpdir / 'stock_cache.xlsx'
            if cache_path.exists():
                drive.upload_file(cache_path.read_bytes(), 'stock_cache.xlsx')

        except Exception as e:
            st.error(f'Erreur lors de l\'import : {e}')
            log(f'ERREUR import SAP : {e}\n{traceback.format_exc()}')
            return

        for line in log_lines:
            log(line)
        log(f'✅ Import SAP terminé — main.xlsx mis à jour par {st.session_state.get("name", "?")}')
        refresh_main_xlsx()
        st.success('Import terminé. main.xlsx mis à jour sur Google Drive.')
        st.rerun()


# ── Étape : Générer les Picking Lists ─────────────────────────────────────────

def step_generate_picking():
    with st.container():
        st.markdown('<div class="sg-card">', unsafe_allow_html=True)
        c1, c2 = st.columns([4, 1])
        with c1:
            st.markdown('**📋 Générer les Picking Lists**')
            st.caption('Allouer le stock (FEFO) et créer les fichiers PL_*.xlsx')
        with c2:
            if st.button('▶ Lancer', key='btn_picking', type='primary'):
                st.session_state['show_picking_config'] = True

        # Panneau de configuration (composants + PO)
        if st.session_state.get('show_picking_config'):
            with st.expander('Configuration', expanded=True):
                st.markdown('**Composants à générer**')
                comp_cols = st.columns(3)
                selected_comps = []
                for i, comp in enumerate(COMPONENTS):
                    with comp_cols[i % 3]:
                        if st.checkbox(comp, value=True, key=f'comp_{comp}'):
                            selected_comps.append(comp)

                st.markdown('**Numéros de PO (optionnel)**')
                po_cols = st.columns(3)
                po_numbers = {}
                for i, comp in enumerate(COMPONENTS):
                    with po_cols[i % 3]:
                        val = st.text_input(f'PO {comp}', key=f'po_{comp}', placeholder='ex: 4500123456')
                        if val:
                            po_numbers[comp] = val

                btn_col1, btn_col2 = st.columns([1, 4])
                with btn_col1:
                    if st.button('Annuler', key='cancel_picking'):
                        st.session_state['show_picking_config'] = False
                        st.rerun()
                with btn_col2:
                    if st.button('▶ Générer', key='confirm_picking', type='primary'):
                        if not selected_comps:
                            st.warning('Sélectionnez au moins un composant.')
                        else:
                            st.session_state['show_picking_config'] = False
                            _do_generate_picking(selected_comps, po_numbers)

        # Fichiers générés disponibles au téléchargement
        if 'generated_picking_files' in st.session_state:
            st.markdown('**Fichiers générés :**')
            for fname, fbytes in st.session_state['generated_picking_files'].items():
                st.download_button(
                    f'⬇ {fname}', data=fbytes, file_name=fname,
                    mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                    key=f'dl_{fname}'
                )

        st.markdown('</div>', unsafe_allow_html=True)


def _do_generate_picking(selected_comps, po_numbers):
    import picking_list_generator as plg_mod
    import sys

    main_bytes = get_main_xlsx()
    if main_bytes is None:
        st.error('main.xlsx introuvable sur Google Drive.')
        return

    with tempfile.TemporaryDirectory() as tmpdir:
        tmpdir = Path(tmpdir)
        main_path = tmpdir / 'main.xlsx'
        main_path.write_bytes(main_bytes)

        # Récupérer pl_counter.json depuis Drive si disponible
        counter_bytes = drive.download_file('pl_counter.json')
        if counter_bytes:
            (tmpdir / 'pl_counter.json').write_bytes(counter_bytes)

        picking_out = tmpdir / 'picking_lists'
        picking_out.mkdir()

        log_lines = []

        class _Capture:
            def write(self, t): log_lines.append(t.rstrip())
            def flush(self): pass

        old_out = sys.stdout
        sys.stdout = _Capture()
        try:
            with st.spinner('Génération des picking lists…'):
                gen = plg_mod.PickingListGenerator(str(main_path))
                gen.load_data()
                gen.remove_staging_locations(target_file=str(main_path))
                gen.generate_picking_lists(components_filter=selected_comps)
                gen.save_picking_lists(output_folder=str(picking_out), shared_dir=tmpdir)
                gen.save_updated_stock(output_file=str(main_path))
        except Exception as e:
            sys.stdout = old_out
            st.error(f'Erreur génération : {e}')
            log(f'ERREUR picking : {e}\n{traceback.format_exc()}')
            return
        finally:
            sys.stdout = old_out

        for line in log_lines:
            log(line)

        # Collecter les fichiers générés
        pl_files = sorted(picking_out.glob('PL_*.xlsx'))
        if not pl_files:
            st.warning('Aucune picking list générée.')
            return

        # Sauvegarder main.xlsx mis à jour sur Drive
        drive.upload_file(main_path.read_bytes(), 'main.xlsx')

        # Sauvegarder pl_counter.json
        counter_path = tmpdir / 'pl_counter.json'
        if counter_path.exists():
            drive.upload_file(counter_path.read_bytes(), 'pl_counter.json')

        # Sauvegarder stock_cache.xlsx
        cache_path = tmpdir / 'stock_cache.xlsx'
        if cache_path.exists():
            drive.upload_file(cache_path.read_bytes(), 'stock_cache.xlsx')

        # Archiver sur Drive dans picking_lists/
        folder_id = drive.get_folder_id()
        pl_folder_id = drive.get_subfolder_id('picking_lists', folder_id)
        generated = {}
        for f in pl_files:
            fbytes = f.read_bytes()
            drive.upload_file(fbytes, f.name, pl_folder_id)
            generated[f.name] = fbytes

        # Sauvegarder PO numbers pour les PowerPoints
        if po_numbers:
            po_bytes = json.dumps(po_numbers, indent=2, ensure_ascii=False).encode()
            drive.upload_file(po_bytes, 'po_numbers.json')

        st.session_state['generated_picking_files'] = generated
        refresh_main_xlsx()
        log(f'✅ {len(pl_files)} picking list(s) générée(s) et uploadées sur Drive.')
        st.success(f'{len(pl_files)} picking list(s) générée(s). Téléchargez-les ci-dessous.')
        st.rerun()


# ── Étape : Mettre à jour les PowerPoints ─────────────────────────────────────

def step_update_pptx():
    with st.container():
        st.markdown('<div class="sg-card">', unsafe_allow_html=True)
        c1, c2 = st.columns([4, 1])
        with c1:
            st.markdown('**📊 Mettre à jour les PowerPoints**')
            st.caption('Remplir les emplacements et N° PO dans les fichiers PW_*.pptx')
        with c2:
            if st.button('▶ Lancer', key='btn_pptx', type='primary'):
                _do_update_pptx()

        if 'generated_pptx_files' in st.session_state:
            st.markdown('**Fichiers générés :**')
            for fname, fbytes in st.session_state['generated_pptx_files'].items():
                st.download_button(
                    f'⬇ {fname}', data=fbytes, file_name=fname,
                    mime='application/vnd.openxmlformats-officedocument.presentationml.presentation',
                    key=f'dl_pptx_{fname}'
                )
        st.markdown('</div>', unsafe_allow_html=True)


def _do_update_pptx():
    import update_all_powerpoints
    import sys

    main_bytes = get_main_xlsx()
    if main_bytes is None:
        st.error('main.xlsx introuvable sur Google Drive.')
        return

    with tempfile.TemporaryDirectory() as tmpdir:
        tmpdir = Path(tmpdir)
        main_path = tmpdir / 'main.xlsx'
        main_path.write_bytes(main_bytes)

        # Télécharger les sources PPTX depuis Drive
        folder_id = drive.get_folder_id()
        sources_folder_id = drive.get_subfolder_id('sources', folder_id)
        sources_dir = tmpdir / 'sources'
        sources_dir.mkdir()

        source_files = drive.list_files(sources_folder_id, pattern='.pptx')
        if not source_files:
            st.error("Aucun fichier source .pptx trouvé dans Drive/sources/.")
            return

        for f in source_files:
            fbytes = drive.download_file(f['name'], sources_folder_id)
            if fbytes:
                (sources_dir / f['name']).write_bytes(fbytes)

        # PO numbers
        po_bytes = drive.download_file('po_numbers.json')
        po_numbers = json.loads(po_bytes.decode()) if po_bytes else {}

        pptx_out = tmpdir / 'powerpoints_updated'
        pptx_out.mkdir()

        log_lines = []

        class _Capture:
            def write(self, t): log_lines.append(t.rstrip())
            def flush(self): pass

        old_out = sys.stdout
        sys.stdout = _Capture()
        try:
            with st.spinner('Mise à jour des PowerPoints…'):
                update_all_powerpoints.main(
                    po_numbers=po_numbers or None,
                    shared_dir=tmpdir,
                )
        except Exception as e:
            sys.stdout = old_out
            st.error(f'Erreur PowerPoints : {e}')
            log(f'ERREUR pptx : {e}\n{traceback.format_exc()}')
            return
        finally:
            sys.stdout = old_out

        for line in log_lines:
            log(line)

        pw_files = sorted(pptx_out.glob('PW_*.pptx'))
        if not pw_files:
            st.warning('Aucun PowerPoint généré.')
            return

        # Upload sur Drive
        pptx_folder_id = drive.get_subfolder_id('powerpoints_updated', folder_id)
        generated = {}
        for f in pw_files:
            fbytes = f.read_bytes()
            drive.upload_file(fbytes, f.name, pptx_folder_id)
            generated[f.name] = fbytes

        # Réinitialiser les PO numbers
        drive.upload_file(b'{}', 'po_numbers.json')

        st.session_state['generated_pptx_files'] = generated
        log(f'✅ {len(pw_files)} PowerPoint(s) mis à jour et uploadés sur Drive.')
        st.success(f'{len(pw_files)} PowerPoint(s) générés. Téléchargez-les ci-dessous.')
        st.rerun()


# ── Étape : Archiver ──────────────────────────────────────────────────────────

def step_archive():
    with st.container():
        st.markdown('<div class="sg-card">', unsafe_allow_html=True)
        c1, c2 = st.columns([4, 1])
        with c1:
            st.markdown('**📦 Archiver les Picking Lists**')
            st.caption('Déplace les PL_*.xlsx du dossier picking_lists vers un dossier d\'archive numéroté')
        with c2:
            if st.button('▶ Archiver', key='btn_archive', type='primary'):
                _do_archive()
        st.markdown('</div>', unsafe_allow_html=True)


def _do_archive():
    folder_id = drive.get_folder_id()
    pl_folder_id = drive.get_subfolder_id('picking_lists', folder_id)
    pl_files = drive.list_files(pl_folder_id, pattern='PL_')
    pl_files = [f for f in pl_files if f['name'].startswith('PL_') and f['name'].endswith('.xlsx')]

    if not pl_files:
        st.info('Aucune picking list à archiver.')
        return

    # Récupérer archive_index.json depuis Drive
    archive_index_bytes = drive.download_file('archive_index.json')
    archive_index = json.loads(archive_index_bytes.decode()) if archive_index_bytes else {'archives': [], 'next_number': 1}
    archive_num = archive_index.get('next_number', 1)
    archive_name = f'archive_{archive_num:03d}_{datetime.now().strftime("%Y%m%d_%H%M")}'

    # Créer le sous-dossier d'archive
    archive_folder_id = drive.get_subfolder_id(archive_name, drive.get_subfolder_id('picking_lists_archive', folder_id))

    service = drive.get_service()
    archived_names = []
    for f in pl_files:
        # Déplacer le fichier (changer le parent)
        service.files().update(
            fileId=f['id'],
            addParents=archive_folder_id,
            removeParents=pl_folder_id,
            fields='id',
        ).execute()
        archived_names.append(f['name'])

    # Mettre à jour archive_index.json
    archive_index['archives'].insert(0, {
        'number': archive_num,
        'folder': archive_name,
        'date': datetime.now().strftime('%d/%m/%Y à %H:%M'),
        'user': st.session_state.get('name', '?'),
        'files': archived_names,
    })
    archive_index['next_number'] = archive_num + 1
    drive.upload_file(json.dumps(archive_index, indent=2).encode(), 'archive_index.json')

    # Vider la session
    st.session_state.pop('generated_picking_files', None)
    log(f'✅ {len(pl_files)} fichier(s) archivé(s) dans {archive_name}.')
    st.success(f'{len(pl_files)} fichier(s) archivé(s).')
    st.rerun()


# ── Étape : Picking Ad-Hoc ────────────────────────────────────────────────────

def step_adhoc():
    with st.container():
        st.markdown('<div class="sg-card">', unsafe_allow_html=True)
        st.markdown('**🔧 Picking Ad-Hoc**')
        st.caption('Génère une picking list à la demande pour des références précises')

        # Récupérer les refs disponibles depuis main.xlsx
        main_bytes = get_main_xlsx()
        all_refs = []
        ref_desc_map = {}
        if main_bytes:
            try:
                for sheet in COMPONENTS:
                    df = pd.read_excel(io.BytesIO(main_bytes), sheet_name=sheet, header=0)
                    for _, row in df.iterrows():
                        ref = str(row.iloc[0]).strip()
                        if ref and ref != 'nan':
                            desc = str(row.iloc[2]).strip() if len(row) > 2 and pd.notna(row.iloc[2]) else ''
                            if desc and desc != 'nan':
                                ref_desc_map[ref] = desc
                            if ref not in all_refs:
                                all_refs.append(ref)
            except Exception:
                pass

        if 'adhoc_items' not in st.session_state:
            st.session_state['adhoc_items'] = []

        # Formulaire d'ajout
        with st.form('adhoc_form', clear_on_submit=True):
            fa, fb, fc, fd = st.columns([3, 1, 1, 1])
            with fa:
                ref_options = [f"{r}  —  {ref_desc_map[r]}" if r in ref_desc_map else r for r in all_refs]
                ref_input = st.selectbox('Référence', options=[''] + ref_options, index=0)
            with fb:
                qty_input = st.number_input('Quantité', min_value=1, value=1, step=1)
            with fc:
                unit_input = st.radio('Unité', ['palette', 'pièce'], horizontal=True)
            with fd:
                st.markdown('<div style="padding-top:1.6rem">', unsafe_allow_html=True)
                add_btn = st.form_submit_button('+ Ajouter')
                st.markdown('</div>', unsafe_allow_html=True)

            if add_btn and ref_input:
                pure_ref = ref_input.split('  —  ')[0].strip().upper()
                st.session_state['adhoc_items'].append({
                    'reference': pure_ref,
                    'quantity': qty_input,
                    'unit': unit_input,
                })

        # Liste des articles
        items = st.session_state['adhoc_items']
        if items:
            df_items = pd.DataFrame(items)
            st.dataframe(df_items, hide_index=True, use_container_width=True)

            col_del, _, col_gen = st.columns([1, 3, 1])
            with col_del:
                if st.button('🗑 Vider la liste', key='adhoc_clear'):
                    st.session_state['adhoc_items'] = []
                    st.rerun()
            with col_gen:
                if st.button('▶ Générer', key='adhoc_generate', type='primary'):
                    _do_adhoc_picking(list(items))
        else:
            st.caption('Aucun article ajouté.')

        if 'adhoc_result' in st.session_state:
            fname, fbytes = st.session_state['adhoc_result']
            st.download_button(
                f'⬇ Télécharger {fname}', data=fbytes, file_name=fname,
                mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                key='dl_adhoc'
            )

        st.markdown('</div>', unsafe_allow_html=True)


def _do_adhoc_picking(items):
    import picking_list_generator as plg_mod
    import sys

    main_bytes = get_main_xlsx()
    if main_bytes is None:
        st.error('main.xlsx introuvable sur Google Drive.')
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

        log_lines = []

        class _Capture:
            def write(self, t): log_lines.append(t.rstrip())
            def flush(self): pass

        old_out = sys.stdout
        sys.stdout = _Capture()
        try:
            with st.spinner('Génération ad-hoc…'):
                gen = plg_mod.PickingListGenerator(str(main_path))
                gen.load_data()
                filename = gen.generate_adhoc_picking_list(
                    items, output_folder=str(picking_out), shared_dir=tmpdir
                )
                gen.save_updated_stock(output_file=str(main_path))
        except Exception as e:
            sys.stdout = old_out
            st.error(f'Erreur ad-hoc : {e}')
            log(f'ERREUR adhoc : {e}\n{traceback.format_exc()}')
            return
        finally:
            sys.stdout = old_out

        for line in log_lines:
            log(line)

        if filename and Path(filename).exists():
            fbytes = Path(filename).read_bytes()
            fname = Path(filename).name

            # Upload sur Drive
            folder_id = drive.get_folder_id()
            pl_folder_id = drive.get_subfolder_id('picking_lists', folder_id)
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
            st.success('Picking list ad-hoc générée.')
            st.rerun()
        else:
            st.warning('Aucun fichier généré.')


# ── Étape : Consulter le stock ────────────────────────────────────────────────

def step_stock():
    with st.container():
        st.markdown('<div class="sg-card">', unsafe_allow_html=True)
        st.markdown('**🔍 Consulter le stock**')
        st.caption('Rechercher une référence et afficher ses emplacements, quantités et dates')

        cache_df = get_stock_cache()

        if cache_df is None:
            st.warning('Aucun cache stock disponible. Lancez d\'abord un import SAP.')
        else:
            st.caption(f'{len(cache_df)} lignes dans le cache')
            search = st.text_input('Référence', placeholder='ex: 1234567-001', key='stock_search')
            if search:
                mask = cache_df['Product'].astype(str).str.upper().str.contains(search.upper())
                result = cache_df[mask]
                if result.empty:
                    st.info('Aucun résultat.')
                else:
                    # Colonnes utiles
                    cols = ['Handling Unit', 'Storage Bin', 'Quantity', 'Base Unit of Measure', 'Shelf Life Expiration Date']
                    cols = [c for c in cols if c in result.columns]
                    total_qty = result['Quantity'].sum() if 'Quantity' in result.columns else None
                    st.dataframe(result[cols], use_container_width=True, hide_index=True)
                    if total_qty is not None:
                        st.metric('Quantité totale', f'{total_qty:g}')

        st.markdown('</div>', unsafe_allow_html=True)


# ── Section BOM ───────────────────────────────────────────────────────────────

def step_bom_sheet():
    with st.container():
        st.markdown('<div class="sg-card">', unsafe_allow_html=True)
        c1, c2 = st.columns([4, 1])
        with c1:
            st.markdown('**📑 Générer l\'onglet BOM**')
            st.caption('Consolider toutes les références dans un onglet dédié de main.xlsx')
        with c2:
            if st.button('▶ Lancer', key='btn_bom_sheet', type='primary'):
                _do_bom_sheet()
        st.markdown('</div>', unsafe_allow_html=True)


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

        log_lines = []

        class _Capture:
            def write(self, t): log_lines.append(t.rstrip())
            def flush(self): pass

        old_out = sys.stdout
        sys.stdout = _Capture()
        try:
            with st.spinner('Génération onglet BOM…'):
                plg_mod.PickingListGenerator(str(main_path)).generate_bom_sheet()
        except Exception as e:
            sys.stdout = old_out
            st.error(f'Erreur BOM sheet : {e}')
            return
        finally:
            sys.stdout = old_out

        for line in log_lines:
            log(line)

        drive.upload_file(main_path.read_bytes(), 'main.xlsx')
        refresh_main_xlsx()
        log('✅ Onglet BOM généré et main.xlsx mis à jour.')
        st.success('Onglet BOM généré. main.xlsx mis à jour sur Drive.')
        st.rerun()


def step_sync_bom():
    with st.container():
        st.markdown('<div class="sg-card">', unsafe_allow_html=True)
        c1, c2, c3 = st.columns([4, 1, 1])
        with c1:
            st.markdown('**🔄 Synchroniser le BOM**')
            st.caption('Propager les changements de main.xlsx vers les PowerPoint sources')
        with c2:
            if st.button('🔍 Simuler', key='btn_sim_bom'):
                _do_sync_bom(dry_run=True)
        with c3:
            if st.button('▶ Synchroniser', key='btn_sync_bom', type='primary'):
                _do_sync_bom(dry_run=False)

        if 'bom_report' in st.session_state:
            fname, fbytes = st.session_state['bom_report']
            st.download_button(f'⬇ Rapport {fname}', data=fbytes, file_name=fname,
                               mime='text/html', key='dl_bom_report')

        st.markdown('</div>', unsafe_allow_html=True)


def _do_sync_bom(dry_run=False):
    import sync_bom
    import sys

    main_bytes = get_main_xlsx()
    if main_bytes is None:
        st.error('main.xlsx introuvable.')
        return

    folder_id = drive.get_folder_id()
    sources_folder_id = drive.get_subfolder_id('sources', folder_id)

    with tempfile.TemporaryDirectory() as tmpdir:
        tmpdir = Path(tmpdir)
        main_path = tmpdir / 'main.xlsx'
        main_path.write_bytes(main_bytes)

        sources_dir = tmpdir / 'sources'
        sources_dir.mkdir()

        source_files = drive.list_files(sources_folder_id)
        for f in source_files:
            fbytes = drive.download_file(f['name'], sources_folder_id)
            if fbytes:
                (sources_dir / f['name']).write_bytes(fbytes)

        # Patch sync_bom pour utiliser tmpdir
        old_excel = sync_bom.EXCEL_FILE
        old_sources = sync_bom.SOURCES_DIR
        sync_bom.EXCEL_FILE  = main_path
        sync_bom.SOURCES_DIR = sources_dir

        log_lines = []

        class _Capture:
            def write(self, t): log_lines.append(t.rstrip())
            def flush(self): pass

        old_out = sys.stdout
        sys.stdout = _Capture()
        try:
            label = 'Simulation' if dry_run else 'Synchronisation'
            with st.spinner(f'{label} BOM…'):
                report_path = sync_bom.main(dry_run=dry_run)
        except Exception as e:
            sys.stdout = old_out
            sync_bom.EXCEL_FILE  = old_excel
            sync_bom.SOURCES_DIR = old_sources
            st.error(f'Erreur sync BOM : {e}')
            log(f'ERREUR sync_bom : {e}\n{traceback.format_exc()}')
            return
        finally:
            sys.stdout = old_out
            sync_bom.EXCEL_FILE  = old_excel
            sync_bom.SOURCES_DIR = old_sources

        for line in log_lines:
            log(line)

        if not dry_run:
            # Uploader les sources modifiées
            for f in sources_dir.glob('*.pptx'):
                drive.upload_file(f.read_bytes(), f.name, sources_folder_id)
            log('✅ Sources PPTX mises à jour sur Drive.')

        if report_path and Path(report_path).exists():
            rname = Path(report_path).name
            rbytes = Path(report_path).read_bytes()
            drive.upload_file(rbytes, rname, folder_id)
            st.session_state['bom_report'] = (rname, rbytes)

        action = 'Simulation terminée' if dry_run else 'Synchronisation terminée'
        st.success(f'{action}. Consultez le rapport ci-dessous.')
        st.rerun()


# ── Workflow complet ──────────────────────────────────────────────────────────

def step_workflow_complet():
    st.markdown("""
    <div style='background:#00B4B4; border-radius:8px; padding:1rem 1.5rem; margin-bottom:1rem;
                color:black; font-weight:700; font-size:1rem; cursor:pointer;'>
        ▶&nbsp;&nbsp;Lancer le workflow complet
    </div>
    """, unsafe_allow_html=True)

    if st.button('▶  Lancer le workflow complet', key='btn_workflow_all', use_container_width=True):
        st.session_state['show_workflow_config'] = True

    if st.session_state.get('show_workflow_config'):
        with st.expander('Configuration du workflow', expanded=True):
            st.markdown('**Composants**')
            comp_cols = st.columns(3)
            selected_comps = []
            for i, comp in enumerate(COMPONENTS):
                with comp_cols[i % 3]:
                    if st.checkbox(comp, value=True, key=f'wf_comp_{comp}'):
                        selected_comps.append(comp)

            st.markdown('**Numéros de PO (optionnel)**')
            po_cols = st.columns(3)
            po_numbers = {}
            for i, comp in enumerate(COMPONENTS):
                with po_cols[i % 3]:
                    val = st.text_input(f'PO {comp}', key=f'wf_po_{comp}', placeholder='ex: 4500123456')
                    if val:
                        po_numbers[comp] = val

            c1, c2 = st.columns([1, 4])
            with c1:
                if st.button('Annuler', key='cancel_workflow'):
                    st.session_state['show_workflow_config'] = False
                    st.rerun()
            with c2:
                if st.button('▶ Lancer le workflow', key='confirm_workflow', type='primary'):
                    if not selected_comps:
                        st.warning('Sélectionnez au moins un composant.')
                    else:
                        st.session_state['show_workflow_config'] = False
                        _do_generate_picking(selected_comps, po_numbers)
                        _do_update_pptx()


# ── Page principale ────────────────────────────────────────────────────────────

def main_app():
    # Header
    user = st.session_state.get('name', '')
    st.markdown(f"""
    <div class="sg-header">
        <p class="sg-teal">BLADE B115</p>
        <h1>Gestion des Picking Lists</h1>
        <p>{VERSION}  &nbsp;·&nbsp; Connecté : {user}</p>
    </div>
    """, unsafe_allow_html=True)

    # Bouton déconnexion dans la sidebar
    with st.sidebar:
        st.markdown(f'**{user}**')
        authenticator.logout('Déconnexion', 'sidebar', key='logout')
        st.divider()
        st.caption(VERSION)

    # Workflow complet
    step_workflow_complet()

    # Section 1 — Workflow de production
    section('Workflow de production')
    step_import_sap()
    step_generate_picking()
    step_update_pptx()
    step_archive()

    # Section 2 — Consultation & Ad-Hoc
    section('Consultation & Ad-Hoc')
    step_adhoc()
    step_stock()

    # Section 3 — Gestion du BOM
    section('Gestion du BOM')
    step_bom_sheet()
    step_sync_bom()

    # Journal
    st.divider()
    section('Journal')
    show_log()


# ── Point d'entrée ─────────────────────────────────────────────────────────────

if st.session_state.get('authentication_status'):
    main_app()
else:
    show_login()
