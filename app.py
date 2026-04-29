import streamlit as st
import tempfile, os
from procesador import generar_reposicion

st.set_page_config(
    page_title="Reposición de Inventario",
    page_icon="📦",
    layout="centered"
)

st.markdown("""
<style>
    .main { max-width: 700px; margin: auto; }
    .stButton>button {
        background-color: #1F3864;
        color: white;
        font-size: 16px;
        padding: 10px 30px;
        border-radius: 8px;
        width: 100%;
    }
    .stButton>button:hover { background-color: #2E75B6; }
    .result-box {
        background: #E2EFDA;
        border-left: 5px solid #1F3864;
        padding: 15px;
        border-radius: 6px;
        margin-top: 10px;
    }
</style>
""", unsafe_allow_html=True)

st.title("📦 Generador de Reposición de Inventario")
st.markdown("Sube tu archivo `.xls` de inventario y descarga el reporte de reposición listo para usar.")

st.divider()

col1, col2 = st.columns([2, 1])
with col1:
    uploaded = st.file_uploader("📁 Archivo de inventario (.xls)", type=["xls", "xlsx"])
with col2:
    dias = st.selectbox("📅 Días de reposición", [30, 45, 60], index=0)

if uploaded:
    st.info(f"✅ Archivo cargado: **{uploaded.name}**")

    if st.button("⚙️ Generar Reposición"):
        with st.spinner("Procesando... por favor espera."):
            try:
                with tempfile.NamedTemporaryFile(delete=False, suffix=".xls") as tmp_in:
                    tmp_in.write(uploaded.read())
                    tmp_path = tmp_in.name

                output_path = tmp_path.replace(".xls", "_reposicion.xlsx")
                out, n_prods, branches = generar_reposicion(tmp_path, dias=dias, output_path=output_path)

                with open(output_path, "rb") as f:
                    xlsx_bytes = f.read()

                os.unlink(tmp_path)
                os.unlink(output_path)

                nombre_salida = uploaded.name.replace(".xls", "").replace(".xlsx", "")
                nombre_salida = f"Reposicion_{nombre_salida}_{dias}dias.xlsx"

                st.markdown(f"""
                <div class="result-box">
                    ✅ <b>Archivo generado correctamente</b><br>
                    📦 <b>{n_prods}</b> productos procesados<br>
                    🏪 <b>{len(branches)}</b> sucursales + Mapa 2 (nueva bodega)<br>
                    📋 <b>5 hojas:</b> Consolidado, Rep+Saldo+Cob, Detalle por Sucursal, Saldos, Sugerido Mapa 2
                </div>
                """, unsafe_allow_html=True)

                st.download_button(
                    label="⬇️ Descargar Reposición.xlsx",
                    data=xlsx_bytes,
                    file_name=nombre_salida,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

            except Exception as e:
                st.error(f"❌ Error al procesar el archivo: {e}")
                raise e

st.divider()
st.markdown("""
**📋 Hojas generadas:**
- **Consolidado Reposición** — Vista resumen con unidades a pedir por sucursal
- **Rep + Saldo + Cob + PVtas** — Detalle en 3 filas: Reponer / Saldo / Cobertura
- **Detalle por Sucursal** — Lista plana con cobertura en días
- **Saldos de Bodega** — Inventario actual por sucursal
- **Sugerido Inventario Mapa 2** — Sugerido para la nueva bodega (30/45/60 días)
""")
