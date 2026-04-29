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
    .main { max-width: 750px; margin: auto; }
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
    .section-title {
        font-weight: bold;
        color: #1F3864;
        font-size: 15px;
        margin-top: 10px;
    }
</style>
""", unsafe_allow_html=True)

st.title("📦 Generador de Reposición de Inventario")
st.markdown("Sube tu archivo `.xls` o `.xlsx` de inventario y descarga el reporte completo.")

st.divider()

# ── Archivo ───────────────────────────────────────────────────────────────────
uploaded = st.file_uploader("📁 Archivo de inventario (.xls / .xlsx)", type=["xls", "xlsx"])

st.divider()

# ── Configuración reposición ──────────────────────────────────────────────────
st.markdown('<p class="section-title">📋 Configuración de Reposición</p>', unsafe_allow_html=True)
col1, col2 = st.columns(2)
with col1:
    dias = st.selectbox("Días de reposición", [30, 45, 60], index=0)
with col2:
    st.markdown("")

# ── Configuración transferencias ──────────────────────────────────────────────
st.markdown('<p class="section-title">🔄 Configuración de Transferencias</p>', unsafe_allow_html=True)
col3, col4 = st.columns(2)
with col3:
    dias_transfer = st.selectbox("Días mínimos a mantener", [30, 60, 90], index=0)
with col4:
    bodega_acopio = st.selectbox("Bodega destino (Libertad / Duran)",
                                  ["06 Mapasingue", "11 Mapa 2"], index=0)

st.divider()

if uploaded:
    st.info(f"✅ Archivo cargado: **{uploaded.name}**")

    if st.button("⚙️ Generar Reporte Completo"):
        with st.spinner("Procesando... por favor espera."):
            try:
                suffix = ".xlsx" if uploaded.name.endswith(".xlsx") else ".xls"
                with tempfile.NamedTemporaryFile(delete=False, suffix=suffix) as tmp_in:
                    tmp_in.write(uploaded.read())
                    tmp_path = tmp_in.name

                output_path = tmp_path + "_reposicion.xlsx"
                out, n_prods, branches = generar_reposicion(
                    tmp_path,
                    dias=dias,
                    output_path=output_path,
                    dias_transfer=dias_transfer,
                    bodega_acopio=bodega_acopio
                )

                with open(output_path, "rb") as f:
                    xlsx_bytes = f.read()

                os.unlink(tmp_path)
                os.unlink(output_path)

                nombre_salida = uploaded.name.rsplit(".", 1)[0]
                nombre_salida = f"Reposicion_{nombre_salida}_{dias}dias.xlsx"

                st.markdown(f"""
                <div class="result-box">
                    ✅ <b>Reporte generado correctamente</b><br>
                    📦 <b>{n_prods}</b> productos procesados<br>
                    🏪 <b>{len(branches)}</b> sucursales + Mapa 2<br>
                    📋 <b>8 hojas:</b> Consolidado · Rep+Saldo · Detalle Expandido ·
                    Detalle Sucursal · Saldos · Mapa 2 · <b>Transferencias</b> · <b>Alertas Atípicas</b>
                </div>
                """, unsafe_allow_html=True)

                st.download_button(
                    label="⬇️ Descargar Reporte.xlsx",
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
- **Consolidado Reposición** — Unidades a pedir por sucursal
- **Rep + Saldo + Cob + PVtas** — 3 filas por producto: Reponer / Saldo / P.Vtas
- **Detalle Expandido** — Columnas horizontales por sucursal
- **Detalle por Sucursal** — Lista plana con cobertura en días
- **Saldos de Bodega** — Inventario actual
- **Sugerido Inventario Mapa 2** — Sugerido nueva bodega
- **Transferencias** — Matriz de transferencias entre sucursales por grupo geográfico
- **Alertas Ventas Atípicas** — Productos con ventas fuera de lo normal ⚠️
""")
