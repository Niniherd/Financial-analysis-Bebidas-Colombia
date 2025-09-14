import pandas as pd
import matplotlib.pyplot as plt
pd.set_option ("display.max_columns", None)
file_path= r"D:\Estados_Financieros_Bebidas_Colombia.xlsx"
er=pd.read_excel(file_path, sheet_name="Estado_Resultados")
bg=pd.read_excel(file_path, sheet_name="Balance_General")

er_long = er.melt(id_vars=["Concepto"],
                        var_name="Año",
                        value_name="Valor_Estado_Resultados")
bg_long = bg.melt(id_vars=["Concepto"],
                        var_name="Año",
                        value_name="Valor_Balance_General")

estados_merged = pd.merge(er_long, bg_long, on="Año", how="outer")
print("\n--- Vista previa del merge por Año ---\n")
print(estados_merged.head(20))

def calculo_de_indicadores(er_long, bg_long):
    indicadores = {}
    for año in er_long["Año"].unique():
        er_ = er_long[er_long["Año"] == año].set_index("Concepto")["Valor_Estado_Resultados"]
        bg_ = bg_long[bg_long["Año"] == año].set_index("Concepto")["Valor_Balance_General"]

        ventas = er_.get("Ventas netas", 0)
        utilidad_bruta = er_.get("Utilidad bruta", 0)
        ebit = er_.get("Utilidad operativa (EBIT)", 0)
        ebt = er_.get("Utilidad antes de impuestos (EBT)", 0)
        utilidad_neta = er_.get("Utilidad neta", 0)

        activos = bg_.get("Total activos", 0)
        pasivos = bg_.get("Total pasivos", 0)
        patrimonio = bg_.get("Total patrimonio", 0)
        pasivo_corriente = bg_.get("Total pasivo corriente", 0)
        activo_corriente = bg_.get("Total activo corriente", 0)
        c_pagar= bg_.get("Cuentas por pagar", 0)
        c_cobrar=bg_.get("Cuentas por cobrar", 0)
        Inventarios=bg_.get("Inventarios", 0)
        costo_ventas=er_.get("Costo de ventas", 0)

        # Guardamos los indicadores en un diccionario
        indicadores[año] = {
            "Margen Bruto": utilidad_bruta / ventas if ventas else None,
            "Margen Operativo": ebit / ventas if ventas else None,
            "Margen EBT": ebt / ventas if ventas else None,
            "Margen Neto": utilidad_neta / ventas if ventas else None,
            "ROE": utilidad_neta / patrimonio if patrimonio else None,
            "Endeudamiento": pasivos / patrimonio if patrimonio else None,
            "Liquidez Corriente": activo_corriente / pasivo_corriente if pasivo_corriente else None,
            "Flujo Caja Bruto": utilidad_neta + er_.get("Depreciaciones", 0),
            "Flujo Caja Libre (aprox)": (utilidad_neta + er_.get("Depreciaciones", 0))
                                        - (bg_.get("Propiedades, planta y equipo", 0)
                                           - bg_.get("Propiedades, planta y equipo", 0)),
            # Eficiencia / Gestión
            "Rotación Activos": ventas / activos if activos else None,
            "Rotación Inventario": costo_ventas / Inventarios if Inventarios else None,
            "Rotación Cartera": ventas / c_cobrar if c_cobrar else None,
            "Rotación Proveedores": costo_ventas / c_pagar if c_pagar else None,
        }

    # Convertimos el diccionario en un DataFrame y transponemos para que los años sean filas
    return pd.DataFrame(indicadores).T
tabla_indicadores = calculo_de_indicadores(er_long, bg_long)

print("\n--- Tabla de Indicadores Financieros ---\n")
print(tabla_indicadores.round(3))  # Redondea a 2 decimales

df_ind = calculo_de_indicadores(er_long, bg_long)

df_ind[["Margen Bruto", "Margen Operativo", "Margen Neto", "ROE"]].plot(marker="o")
plt.title("Indicadores de Rentabilidad")
plt.ylabel("Ratio")
plt.grid(True)
plt.show()

# Liquidez
df_ind[["Liquidez Corriente"]].plot(marker="o", color="teal")
plt.title("Indicador de Liquidez Corriente")
plt.ylabel("Veces")
plt.grid(True)
plt.show()

# Endeudamiento
df_ind[["Endeudamiento"]].plot(marker="o", color="red")
plt.title("Nivel de Endeudamiento")
plt.ylabel("Pasivos / Patrimonio")
plt.grid(True)
plt.show()

# Eficiencia
df_ind[["Rotación Activos"]].plot(marker="o", color="purple")
plt.title("Rotación de Activos")
plt.ylabel("Ventas / Activos")
plt.grid(True)
plt.show()