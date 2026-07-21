# 🥚 Dashboard de Control Ejecutivo - Costo del Huevo Fértil

Aplicación interactiva desarrollada por Juan Salgado para el análisis, control de variación, simulaciones de escenarios (What-If) y proyecciones futuras del costo de producción de huevo fértil en la operación

---

## 🚀 Características Principales

El dashboard está dividido en **4 módulos estratégicos**:

1. **🔍 Análisis de Variación Intermensual (VIF / Waterfall):** 
   * Descompone la diferencia de costo por huevo fértil ($\$/huevo$) entre dos períodos seleccionados.
   * Explica en detalle qué rubro (Alimento, Depreciación, Arriendo, Mano de Obra, etc.) impulsó el aumento o la disminución del costo.
2. **📊 Impacto por Ítem:**
   * Visualiza la estructura de costos del mes mediante gráficos de participación (%) y barras de costo unitario.
   * Identifica rápidamente el Pareto del gasto.
3. **🎛️ Simulador de Escenarios y Margen (What-If):**
   * Ajuste de variables clave en tiempo real:
     * Precio del alimento ($\$/Kg$)
     * Consumo de alimento ($g/huevo$)
     * Porcentaje de fertilidad ($\%$)
     * Precio de venta objetivo ($\$/huevo$)
   * Recalcula automáticamente el nuevo costo unitario, el margen ($/huevo) y la utilidad proyectada.
4. **🔮 Proyección Futura:**
   * Modelo de suavizado exponencial (Holt-Winters) para proyectar el costo por huevo fértil en los próximos $1$ a $6$ meses.

---
