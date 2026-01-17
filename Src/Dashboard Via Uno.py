import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from datetime import datetime
import warnings
warnings.filterwarnings("ignore")

# Configuración de la página
st.set_page_config(
    page_title="Dashboard Ventas Bsale",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Estilos CSS personalizados
st.markdown("""
    <style>
    /* Ajuste del ancho principal */
    .main { padding: 0rem 1rem; }

    /* Contenedor de las métricas (varios selectores por compatibilidad) */
    [data-testid="stMetric"], [data-testid="metric-container"], div[data-testid="stMetric"] {
        background-color: #0f1724 !important;  /* fondo oscuro (ajusta aquí) */
        padding: 12px 14px !important;
        border-radius: 10px !important;
        box-shadow: 0 4px 10px rgba(2,6,23,0.35) !important;
        text-align: center;
    }

    /* Valor grande (número) */
    [data-testid="stMetricValue"] {
        color: #ffffff !important;
        font-size: 26px !important;
        font-weight: 600 !important;
    }

    /* Label / título */
    [data-testid="stMetricLabel"] {
        color: #c6d0df !important;
        font-size: 13px !important;
    }

    /* Si aparece delta/extra (por si) */
    [data-testid="stMetric"] > div > div > div:nth-child(3),
    .stMetricDelta {
        color: #9ae6b4 !important;
    }

    /* Encabezados */
    h1 { color: #1F4E78; }

    /* Mantener tu estilo de footer/otros */
    .footer { text-align: center; color: #666; padding: 20px 0; }

    </style>
""", unsafe_allow_html=True)

# Carga de datos con caché
@st.cache_data
def cargar_excel(archivo):
    """Carga el Excel procesado y convierte fechas"""
    df = pd.read_excel(archivo)
    
    # Convertir fecha
    df["Fecha de Emisión"] = pd.to_datetime(
        df["Fecha de Emisión"],
        dayfirst=True,
        errors="coerce"
    )
    
    return df

# Interfaz principal
st.title("📊 Dashboard de Ventas Via uno")
st.markdown("---")

# Carga de archivo en sidebar
with st.sidebar:
    st.header("📁 Cargar Datos")
    
    archivo = st.file_uploader(
        "Ventas Automatizado (Excel procesado)",
        type=["xlsx"],
        help="Sube el archivo 'Ventas Automatizado.xlsx' generado por el notebook"
    )
    
    st.markdown("---")
    
    # Opción de incluir devoluciones
    incluir_devoluciones = st.checkbox(
        "Incluir devoluciones (notas de crédito)",
        value=False,
        help="Marca esta opción para incluir las notas de crédito en el análisis"
    )

if archivo:
    
    with st.spinner('Cargando datos...'):
        df = cargar_excel(archivo)
    
    if not incluir_devoluciones:
        df = df[df["Tipo Movimiento"].str.lower() == "venta"].copy()
    
    # Filtros en sidebar
    with st.sidebar:
        st.subheader("🔍 Filtros")

        # --- Fechas ---
        if df["Fecha de Emisión"].notna().any():
            fecha_min = df["Fecha de Emisión"].min()
            fecha_max = df["Fecha de Emisión"].max()

            fecha_inicio, fecha_fin = st.date_input(
                "Rango de fechas",
                value=(fecha_min, fecha_max),
                min_value=fecha_min,
                max_value=fecha_max
            )
        else:
            st.warning("No hay fechas válidas en los datos")
            fecha_inicio = fecha_fin = None

        # Marketplace
        marketplaces_disponibles = df["Marketplace"].dropna().unique()
        marketplace = st.selectbox(
            "Marketplace",
            ["Todos"] + sorted(list(map(str, marketplaces_disponibles)))
        )

        # Sucursal
        sucursales_disponibles = df["Sucursal"].dropna().unique()
        sucursal = st.selectbox(
            "Sucursal",
            ["Todas"] + sorted(list(map(str, sucursales_disponibles)))
        )

        # Filtro específico: Tipo de Producto / Servicio
        tipos_disponibles = df["Tipo de Producto / Servicio"].dropna().unique()
        tipo_producto = st.multiselect(
            "Tipo de Producto / Servicio (selecciona para incluir)",
            options=sorted(list(map(str, tipos_disponibles))),
            default=[],
            help="Dejar vacío = no filtrar por este campo"
        )

        # Filtro genérico elegir columna y valores
        st.markdown("---")
        st.markdown("#### 🔧 Filtro avanzado (cualquier columna)")
        # Elegir columnas candidatas: pocas categorias o texto
        # Puedes ajustar max_unique si quieres permitir más/menos columnas
        max_unique = 500
        columnas_candidatas = [
            col for col in df.columns
            if df[col].nunique(dropna=True) <= max_unique
        ]
        # Orden y poner aviso si muchas columnas no aptas
        if not columnas_candidatas:
            st.info("No hay columnas con pocos valores únicos para el filtro avanzado.")
            columna_avanzada = "Ninguna"
        else:
            columna_avanzada = st.selectbox(
                "Columna a filtrar (avanzado)",
                ["Ninguna"] + sorted(columnas_candidatas)
            )

        if columna_avanzada and columna_avanzada != "Ninguna":
            valores_unicos = sorted(df[columna_avanzada].dropna().astype(str).unique())
            st.caption(f"{len(valores_unicos)} valores únicos en '{columna_avanzada}'")
            valores_seleccionados = st.multiselect(
                f"Selecciona valores de '{columna_avanzada}'",
                options=valores_unicos,
                default=[],
                help="Usa la búsqueda para encontrar valores rápido"
            )
            modo = st.radio(
                "Modo",
                options=["Incluir solo seleccionados", "Excluir seleccionados"],
                index=0,
                horizontal=True
            )
            aplicar_columna_avanzada = st.checkbox("Aplicar filtro avanzado", value=True)
        else:
            valores_seleccionados = []
            modo = "Incluir solo seleccionados"
            aplicar_columna_avanzada = False

        
    # Filtrar DataFrame según selecciones
    df_filtro = df.copy()

    if fecha_inicio and fecha_fin:
        df_filtro = df_filtro[
            (df_filtro["Fecha de Emisión"] >= pd.Timestamp(fecha_inicio)) &
            (df_filtro["Fecha de Emisión"] <= pd.Timestamp(fecha_fin))
        ]

    if marketplace != "Todos":
        df_filtro = df_filtro[df_filtro["Marketplace"] == marketplace]

    if sucursal != "Todas":
        df_filtro = df_filtro[df_filtro["Sucursal"] == sucursal]

    # Aplicar filtro específico de Tipo de Producto / Servicio (si el usuario seleccionó valores)
    if tipo_producto:
        df_filtro = df_filtro[df_filtro["Tipo de Producto / Servicio"].astype(str).isin(tipo_producto)]

    # Aplicar filtro avanzado (columna elegida por el usuario)
    if aplicar_columna_avanzada and columna_avanzada and columna_avanzada != "Ninguna" and valores_seleccionados:
        if modo == "Incluir solo seleccionados":
            df_filtro = df_filtro[df_filtro[columna_avanzada].astype(str).isin(valores_seleccionados)]
        else:  # Excluir seleccionados
            df_filtro = df_filtro[~df_filtro[columna_avanzada].astype(str).isin(valores_seleccionados)]

    
    # Validacion: si no hay datos tras filtros
    if df_filtro.empty:
        st.warning("⚠️ No hay datos con los filtros seleccionados")
        st.stop()
    
    # KPIs principales
    st.header("📈 Métricas Principales")
    
    ventas = df_filtro["Venta Total Neta"].sum()
    costos = abs(df_filtro["Costo Total Neto"].sum())  # Valor absoluto
    margen = df_filtro["Margen"].sum()
    margen_pct = (margen / ventas * 100) if ventas != 0 else 0
    unidades = df_filtro["Cantidad"].sum()
    
    c1, c2, c3, c4, c5 = st.columns(5)
    
    with c1:
        st.metric("💰 Ventas Netas", f"${ventas:,.0f}")
    
    with c2:
        st.metric("💸 Costos", f"${costos:,.0f}")
    
    with c3:
        st.metric("📊 Margen", f"${margen:,.0f}")
    
    with c4:
        st.metric("📈 Margen %", f"{margen_pct:.2f}%")
    
    with c5:
        st.metric("📦 Unidades", f"{int(unidades):,}")
    
    st.markdown("---")
    
    # Tablas y gráficos en pestañas
    tab1, tab2, tab3, tab4, tab5 = st.tabs([
        "📊 Resumen General",
        "🛍️ Marketplace",
        "📦 Tipo de Producto",
        "🏆 Rankings",
        "📅 Datos Detallados"
    ])
    
    with tab1:
        st.subheader("Evolución de Ventas en el Tiempo")
        
        # Opciones de periodicidad y visualización
        col_config1, col_config2 = st.columns([1, 3])
        
        with col_config1:
            periodicidad = st.selectbox(
                "Periodicidad",
                options=["Diario", "Semanal", "Mensual"],
                index=2
            )
        
        with col_config2:
            col_a, col_b = st.columns(2)
            with col_a:
                mostrar_margen = st.checkbox("Mostrar línea de Margen", value=True)
            with col_b:
                ma_window = st.slider("Media móvil (periodos)", 0, 30, 0)
        
        # Preparar datos según periodicidad
        if df_filtro['Fecha de Emisión'].notna().any():
            tmp = df_filtro.copy()
            
            if periodicidad == "Diario":
                tmp['Periodo'] = tmp['Fecha de Emisión'].dt.date
            elif periodicidad == "Semanal":
                tmp['Periodo'] = tmp['Fecha de Emisión'].dt.to_period('W').apply(
                    lambda r: r.start_time.date()
                )
            else:  # Mensual
                tmp['Periodo'] = tmp['Fecha de Emisión'].dt.to_period('M').apply(
                    lambda r: r.start_time.date()
                )
            
            tendencia = tmp.groupby('Periodo').agg(
                Ventas=('Venta Total Neta', 'sum'),
                Margen=('Margen', 'sum'),
                Unidades=('Cantidad', 'sum')
            ).reset_index().sort_values('Periodo')
            
            # Gráfico de líneas
            fig_t = go.Figure()
            
            fig_t.add_trace(go.Scatter(
                x=tendencia['Periodo'],
                y=tendencia['Ventas'],
                name='Ventas',
                mode='lines+markers',
                line=dict(color='#1F4E78', width=3),
                marker=dict(size=8),
                hovertemplate='<b>%{x}</b><br>Ventas: $%{y:,.0f}<extra></extra>'
            ))
            
            if mostrar_margen:
                fig_t.add_trace(go.Scatter(
                    x=tendencia['Periodo'],
                    y=tendencia['Margen'],
                    name='Margen',
                    mode='lines+markers',
                    line=dict(color='#2ECC71', width=3),
                    marker=dict(size=8),
                    yaxis="y2",
                    hovertemplate='<b>%{x}</b><br>Margen: $%{y:,.0f}<extra></extra>'
                ))
            
            # Media móvil
            if ma_window and ma_window > 0:
                tmp_ma = tendencia[['Periodo', 'Ventas']].set_index('Periodo').sort_index()
                tmp_ma['MA'] = tmp_ma['Ventas'].rolling(window=ma_window, min_periods=1).mean()
                
                fig_t.add_trace(go.Scatter(
                    x=tmp_ma.index,
                    y=tmp_ma['MA'],
                    name=f'MA {ma_window}',
                    mode='lines',
                    line=dict(color='#E67E22', width=2, dash='dash'),
                    hovertemplate='<b>%{x}</b><br>MA: $%{y:,.0f}<extra></extra>'
                ))
            
            # Layout
            fig_t.update_layout(
                title="Ventas y Margen en el Tiempo",
                xaxis_title="Periodo",
                yaxis_title="Ventas ($)",
                hovermode="x unified",
                height=500,
                legend=dict(
                    orientation="h",
                    yanchor="bottom",
                    y=1.02,
                    xanchor="right",
                    x=1
                )
            )
            
            if mostrar_margen:
                fig_t.update_layout(
                    yaxis2=dict(
                        title="Margen ($)",
                        overlaying="y",
                        side="right"
                    )
                )
            
            st.plotly_chart(fig_t, use_container_width=True)
            
            # Tabla resumen
            with st.expander("📋 Ver tabla de tendencia"):
                st.dataframe(
                    tendencia.style.format({
                        'Ventas': '${:,.0f}',
                        'Margen': '${:,.0f}',
                        'Unidades': '{:,.0f}'
                    }),
                    use_container_width=True
                )
        else:
            st.info("No hay fechas válidas para mostrar tendencias")
    
    with tab2:
        st.subheader("Ventas y Margen por Marketplace")
        
        resumen_mp = df_filtro.groupby("Marketplace").agg(
            Ventas=("Venta Total Neta", "sum"),
            Costos=("Costo Total Neto", "sum"),
            Margen=("Margen", "sum"),
            Documentos=("Numero Documento", "nunique"),
            Unidades=("Cantidad", "sum")
        ).reset_index()
        
        resumen_mp["Costos"] = abs(resumen_mp["Costos"])
        resumen_mp["Margen %"] = (resumen_mp["Margen"] / resumen_mp["Ventas"] * 100).round(2)
        resumen_mp = resumen_mp.sort_values("Ventas", ascending=False)
        
        # Gráfico de barras
        fig_mp = px.bar(
            resumen_mp,
            x="Marketplace",
            y="Ventas",
            color="Margen %",
            title="Ventas por Marketplace",
            color_continuous_scale="RdYlGn",
            text="Ventas"
        )
        fig_mp.update_traces(
            texttemplate="$%{text:,.0f}",
            textposition="outside"
        )
        fig_mp.update_layout(height=500)
        
        st.plotly_chart(fig_mp, use_container_width=True)
        
        # Tabla
        st.dataframe(
            resumen_mp.style.format({
                "Ventas": "${:,.0f}",
                "Costos": "${:,.0f}",
                "Margen": "${:,.0f}",
                "Margen %": "{:.2f}%",
                "Documentos": "{:,.0f}",
                "Unidades": "{:,.0f}"
            }),
            use_container_width=True
        )
    
    with tab3:
        st.subheader("Ventas y Margen por Tipo de Producto / Servicio")
        
        resumen_tipo = df_filtro.groupby("Tipo de Producto / Servicio").agg(
            Ventas=("Venta Total Neta", "sum"),
            Costos=("Costo Total Neto", "sum"),
            Margen=("Margen", "sum"),
            Cantidad=("Cantidad", "sum")
        ).reset_index()
        
        resumen_tipo["Costos"] = abs(resumen_tipo["Costos"])
        resumen_tipo["Margen %"] = (resumen_tipo["Margen"] / resumen_tipo["Ventas"] * 100).round(2)
        resumen_tipo = resumen_tipo.sort_values("Ventas", ascending=False)
        
        # Gráfico de barras
        fig_tipo = px.bar(
            resumen_tipo,
            x="Tipo de Producto / Servicio",
            y="Ventas",
            color="Margen %",
            title="Ventas por Tipo de Producto",
            color_continuous_scale="RdYlGn",
            text="Ventas"
        )
        fig_tipo.update_traces(
            texttemplate="$%{text:,.0f}",
            textposition="outside"
        )
        fig_tipo.update_layout(height=500)
        
        st.plotly_chart(fig_tipo, use_container_width=True)
        
        # Tabla
        st.dataframe(
            resumen_tipo.style.format({
                "Ventas": "${:,.0f}",
                "Costos": "${:,.0f}",
                "Margen": "${:,.0f}",
                "Margen %": "{:.2f}%",
                "Cantidad": "{:,.0f}"
            }),
            use_container_width=True
        )
    
    with tab4:
        st.subheader("🏆 Rankings")
        
        ranking = df_filtro.groupby("Tipo de Producto / Servicio").agg(
            Ventas=("Venta Total Neta", "sum"),
            Costos=("Costo Total Neto", "sum"),
            Margen=("Margen", "sum"),
            Cantidad=("Cantidad", "sum")
        ).reset_index()
        
        ranking["Costos"] = abs(ranking["Costos"])
        ranking["Margen %"] = (ranking["Margen"] / ranking["Ventas"] * 100).round(2)
        
        col1, col2 = st.columns(2)
        
        with col1:
            st.markdown("### 🔝 Top 10 por Ventas")
            top_ventas = ranking.sort_values("Ventas", ascending=False).head(10)
            
            st.dataframe(
                top_ventas.style.format({
                    "Ventas": "${:,.0f}",
                    "Costos": "${:,.0f}",
                    "Margen": "${:,.0f}",
                    "Margen %": "{:.2f}%",
                    "Cantidad": "{:,.0f}"
                }),
                use_container_width=True
            )
        
        with col2:
            st.markdown("### ⚠️ Bottom 10 por Margen")
            bottom_margen = ranking.sort_values("Margen").head(10)
            
            st.dataframe(
                bottom_margen.style.format({
                    "Ventas": "${:,.0f}",
                    "Costos": "${:,.0f}",
                    "Margen": "${:,.0f}",
                    "Margen %": "{:.2f}%",
                    "Cantidad": "{:,.0f}"
                }),
                use_container_width=True
            )
    
    with tab5:
        st.subheader("📋 Explorador de Datos")
        
        st.info(f"Mostrando {len(df_filtro):,} registros de {len(df):,} totales")
        
        # Selector de columnas
        columnas_disponibles = df_filtro.columns.tolist()
        columnas_mostrar = st.multiselect(
            "Selecciona columnas a mostrar",
            columnas_disponibles,
            default=['Fecha de Emisión', 'Tipo de Documento', 'Numero Documento', 
                    'Marketplace', 'Sucursal', 'Producto / Servicio', 
                    'Cantidad', 'Venta Total Neta', 'Margen', '% Margen']
        )
        
        if columnas_mostrar:
            st.dataframe(
                df_filtro[columnas_mostrar],
                use_container_width=True,
                height=400
            )
            
            # Botón de descarga
            csv = df_filtro[columnas_mostrar].to_csv(index=False).encode('utf-8')
            st.download_button(
                label="📥 Descargar datos filtrados (CSV)",
                data=csv,
                file_name=f"ventas_filtradas_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv",
                mime="text/csv"
            )
        else:
            st.warning("Selecciona al menos una columna para mostrar")

else:
    # Pantalla inicial sin archivo
    st.info("👋 Bienvenido al Dashboard de Ventas Via Uno")
    
    st.markdown("""
    ### 📋 Instrucciones

    1. **Carga el archivo** usando el botón en el panel lateral

    2. **Explora tus datos** con:
       - 📊 Resumen temporal
       - 🛍️ Análisis por marketplace
       - 📦 Análisis por producto
       - 🏆 Rankings de rendimiento
       - 📅 Vista detallada de datos

    """)
    
    with st.expander("ℹ️ Estructura del archivo esperado"):
        st.markdown("""
        El archivo **debe contener** las siguientes columnas:
        
        - `Fecha de Emisión`
        - `Tipo Movimiento` (venta / devolución)
        - `Tipo de Documento`
        - `Numero Documento`
        - `Marketplace`
        - `Sucursal`
        - `Tipo de Producto / Servicio`
        - `Producto / Servicio`
        - `Cantidad`
        - `Venta Total Neta`
        - `Costo Total Neto`
        - `Margen`
        - `% Margen`
        """)

st.markdown("---")
st.markdown("""
<div style='text-align: center; color: #666; padding: 20px 0;'>
    <p><b>Dashboard de Ventas Via Uno</b> | Desarrollado con Streamlit</p>
</div>
""", unsafe_allow_html=True)