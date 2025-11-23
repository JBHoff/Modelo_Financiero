#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Genera un Excel completamente armado con:
- Escenarios (Optimista / Base / Pesimista)
- Tablas de sensibilidad
- Dashboard con KPIs y gráficos
- CSVs y ZIP con todo

Ejecutar con Python 3.8+, requiere: pandas, xlsxwriter, openpyxl, numpy
"""

import pandas as pd
import numpy as np
import xlsxwriter
import zipfile
from io import BytesIO
import math

# -----------------------------
# FUNCIONES FINANCIERAS
# -----------------------------
def npv(rate, cashflows):
    """Net Present Value: cashflows as list with year0..n"""
    rate = float(rate)
    return sum([cf / ((1 + rate) ** t) for t, cf in enumerate(cashflows)])

def irr(cashflows, guess=0.1):
    """IRR via numpy.roots on polynomial. Returns real root if found, else None."""
    # Coefficients for polynomial: cf0 + cf1*(1+r)^-1 + ... -> convert to polynomial by multiplying by (1+r)^n
    # Equivalent polynomial: sum_{t=0..n} cf_t * x^{n-t} = 0, where x = 1 + r
    cf = list(cashflows)
    # If all flows are same sign IRR undefined
    if all([c >= 0 for c in cf]) or all([c <= 0 for c in cf]):
        return None
    coeffs = cf[:]  # cf0..cfn
    # polynomial coefficients highest power first
    poly = coeffs.copy()
    # find roots
    roots = np.roots(poly)
    # select real root > 0 (x = 1+r) => r > -1
    real_rs = []
    for root in roots:
        if np.isreal(root):
            x = float(np.real(root))
            r = x - 1.0
            if r > -0.999999:
                real_rs.append(r)
    if not real_rs:
        # fallback: try secant/newton numeric method
        try:
            r = guess
            for i in range(200):
                f = npv(r, cf)
                # derivative approx
                d = sum([-t * cf[t] / ((1 + r) ** (t + 1)) for t in range(1, len(cf))])
                if d == 0:
                    break
                r_new = r - f / d
                if abs(r_new - r) < 1e-9:
                    r = r_new
                    break
                r = r_new
            return r
        except Exception:
            return None
    # choose root closest to guess
    if real_rs:
        # pick root in reasonable bounds
        candidates = [r for r in real_rs if r > -0.9999 and r < 10]
        if not candidates:
            candidates = real_rs
        # choose median
        return sorted(candidates, key=lambda x: abs(x - guess))[0]
    return None

# -----------------------------
# 1) SUPUESTOS BASE
# -----------------------------
supuestos = {
    "Cantidad de montacargas": 15,
    "Horas de uso por día": 15.5,
    "Días por semana": 6,
    "Costo por montacargas (USD)": 280000,
    "Tipo de cambio (MXN/USD)": 18,
    "Renta mensual por unidad (MXN)": 50000,
    "Leasing tasa anual (estimada)": 0.14,
    "Leasing plazo meses": 36,
    "Valor residual (%)": 0.10,
    "Tasa descuento (VPN) - Base": 0.12,
    "Mantenimiento anual (% del valor) - Base": 0.05,
    "Reemplazo batería (% del valor) - Base": 0.25,
    "Horizonte análisis (años)": 10
}

# Cálculos derivados
units = int(supuestos["Cantidad de montacargas"])
cost_unit_mxn = supuestos["Costo por montacargas (USD)"] * supuestos["Tipo de cambio (MXN/USD)"]
total_purchase = cost_unit_mxn * units
rent_monthly_fleet = supuestos["Renta mensual por unidad (MXN)"] * units
# leasing monthly fleet estimate provided earlier (we keep as calculation with annual rate)
lease_monthly_est_unit = 155900  # kept from prior estimate; script will also compute via annuity approx
lease_monthly_fleet = lease_monthly_est_unit * units
residual_unit = cost_unit_mxn * supuestos["Valor residual (%)"]
residual_total = residual_unit * units
horizon = int(supuestos["Horizonte análisis (años)"])

# -----------------------------
# 2) ESCENARIOS (Multipliers)
# -----------------------------
# We create three scenarios: Optimista, Base, Pesimista
scenarios = {
    "Optimista": {
        "tipo_cambio": supuestos["Tipo de cambio (MXN/USD)"] * 0.98,
        "costo_unit_usd": supuestos["Costo por montacargas (USD)"] * 0.95,
        "renta_mensual_unit": supuestos["Renta mensual por unidad (MXN)"] * 0.9,
        "tasa_descuento": supuestos["Tasa descuento (VPN) - Base"] * 0.9,
        "mantenimiento_pct": supuestos["Mantenimiento anual (% del valor) - Base"] * 0.9,
        "battery_pct": supuestos["Reemplazo batería (% del valor) - Base"] * 0.9,
        "leasing_tasa": supuestos["Leasing tasa anual (estimada)"] * 0.95
    },
    "Base": {
        "tipo_cambio": supuestos["Tipo de cambio (MXN/USD)"],
        "costo_unit_usd": supuestos["Costo por montacargas (USD)"],
        "renta_mensual_unit": supuestos["Renta mensual por unidad (MXN)"],
        "tasa_descuento": supuestos["Tasa descuento (VPN) - Base"],
        "mantenimiento_pct": supuestos["Mantenimiento anual (% del valor) - Base"],
        "battery_pct": supuestos["Reemplazo batería (% del valor) - Base"],
        "leasing_tasa": supuestos["Leasing tasa anual (estimada)"]
    },
    "Pesimista": {
        "tipo_cambio": supuestos["Tipo de cambio (MXN/USD)"] * 1.06,
        "costo_unit_usd": supuestos["Costo por montacargas (USD)"] * 1.08,
        "renta_mensual_unit": supuestos["Renta mensual por unidad (MXN)"] * 1.15,
        "tasa_descuento": supuestos["Tasa descuento (VPN) - Base"] * 1.25,
        "mantenimiento_pct": supuestos["Mantenimiento anual (% del valor) - Base"] * 1.25,
        "battery_pct": supuestos["Reemplazo batería (% del valor) - Base"] * 1.25,
        "leasing_tasa": supuestos["Leasing tasa anual (estimada)"] * 1.2
    }
}

# -----------------------------
# 3) GENERAR FLUJOS POR ESCENARIO Y ALTERNATIVA
# -----------------------------
years = list(range(0, horizon + 1))  # 0..10 inclusive

def generate_flows_for_params(cost_unit_usd, tipo_cambio, renta_unit_monthly, maint_pct, battery_pct, lease_monthly_fleet_local, residual_total_local):
    """Returns three lists of cashflows (length horizon+1) for compra, renta, leasing."""
    cost_unit_mxn_local = cost_unit_usd * tipo_cambio
    total_purchase_local = cost_unit_mxn_local * units
    # Compra
    compra = []
    for y in years:
        if y == 0:
            inv = - total_purchase_local
        else:
            inv = 0
        maint = - (total_purchase_local * maint_pct) if y >= 1 else 0
        batt = - (total_purchase_local * battery_pct) if y == 5 else 0
        compra.append(inv + maint + batt)
    # Renta: monthly *12
    renta = []
    renta_annual = - (renta_unit_monthly * units) * 12
    for y in years:
        renta.append(0 if y == 0 else renta_annual)
    # Leasing: using lease_monthly_fleet_local (monthly total fleet)
    leasing = []
    lease_annual = - lease_monthly_fleet_local * 12
    for y in years:
        pay = lease_annual if 1 <= y <= 3 else 0
        resid = - residual_total_local if y == 3 else 0
        leasing.append(pay + resid)
    return compra, renta, leasing

# We'll need an estimate for lease_monthly_fleet per scenario.
# Use the prior estimated monthly per unit as baseline, but adjust if scenario changes cost or leasing rate.
def estimate_lease_monthly_unit(cost_unit_mxn_local, annual_rate, months=36, residual_pct=0.10):
    """Estimate monthly payment (annuity) for financed amount = price - residual discounted at annual_rate."""
    price = cost_unit_mxn_local
    residual = price * residual_pct
    financed = price - residual
    # monthly rate
    r_month = (1 + annual_rate) ** (1/12) - 1
    if r_month == 0:
        return financed / months
    ann = financed * (r_month * (1 + r_month) ** months) / ((1 + r_month) ** months - 1)
    return ann

# Create storage
scenario_results = {}

for sname, params in scenarios.items():
    # estimate lease monthly unit with scenario cost and scenario leasing rate
    cost_unit_mxn_local = params["costo_unit_usd"] * params["tipo_cambio"]
    lease_monthly_unit_est = estimate_lease_monthly_unit(cost_unit_mxn_local, params["leasing_tasa"], months=36, residual_pct=0.10)
    lease_monthly_fleet_local = lease_monthly_unit_est * units
    compra_cf, renta_cf, leasing_cf = generate_flows_for_params(
        params["costo_unit_usd"],
        params["tipo_cambio"],
        params["renta_mensual_unit"],
        params["mantenimiento_pct"],
        params["battery_pct"],
        lease_monthly_fleet_local,
        residual_total_local = (params["costo_unit_usd"] * params["tipo_cambio"]) * units * 0.10
    )
    # compute NPV and IRR using Python functions (for reporting)
    npv_compra = npv(params["tasa_descuento"], compra_cf)
    npv_renta = npv(params["tasa_descuento"], renta_cf)
    npv_leasing = npv(params["tasa_descuento"], leasing_cf)
    irr_compra = irr(compra_cf)
    irr_renta = irr(renta_cf)
    irr_leasing = irr(leasing_cf)
    scenario_results[sname] = {
        "params": params,
        "lease_monthly_unit_est": lease_monthly_unit_est,
        "lease_monthly_fleet_est": lease_monthly_fleet_local,
        "compra_cf": compra_cf,
        "renta_cf": renta_cf,
        "leasing_cf": leasing_cf,
        "npv_compra": npv_compra,
        "npv_renta": npv_renta,
        "npv_leasing": npv_leasing,
        "irr_compra": irr_compra,
        "irr_renta": irr_renta,
        "irr_leasing": irr_leasing
    }

# -----------------------------
# 4) CREAR DATAFRAMES Y CSVS
# -----------------------------
# Supuestos DF
df_supuestos = pd.DataFrame.from_dict(supuestos, orient='index', columns=["Valor"]).reset_index().rename(columns={"index":"Variable"})

# Flujos por alternativa y escenario: we'll create multi-index table
rows = []
for sname, res in scenario_results.items():
    for t, year in enumerate(years):
        rows.append({
            "Escenario": sname,
            "Año": year,
            "Compra_Flujo": res["compra_cf"][t],
            "Renta_Flujo": res["renta_cf"][t],
            "Leasing_Flujo": res["leasing_cf"][t],
        })
df_flows = pd.DataFrame(rows)

# Graf (acumulados)
grows = []
for sname, res in scenario_results.items():
    acc_c = np.cumsum(res["compra_cf"])
    acc_r = np.cumsum(res["renta_cf"])
    acc_l = np.cumsum(res["leasing_cf"])
    for t in range(len(years)):
        grows.append({
            "Escenario": sname,
            "Año": t,
            "Compra_Acumulado": acc_c[t],
            "Renta_Acumulado": acc_r[t],
            "Leasing_Acumulado": acc_l[t]
        })
df_graf = pd.DataFrame(grows)

# Resumen NPVs/IRRs
summary_rows = []
for sname, res in scenario_results.items():
    summary_rows.append({
        "Escenario": sname,
        "NPV_Compra": res["npv_compra"],
        "NPV_Renta": res["npv_renta"],
        "NPV_Leasing": res["npv_leasing"],
        "IRR_Compra": res["irr_compra"],
        "IRR_Renta": res["irr_renta"],
        "IRR_Leasing": res["irr_leasing"],
        "Lease_Monthly_Unit": res["lease_monthly_unit_est"],
        "Lease_Monthly_Fleet": res["lease_monthly_fleet_est"]
    })
df_summary = pd.DataFrame(summary_rows)

# -----------------------------
# 5) SENSIBILIDAD
# -----------------------------
# Sensitivity 1: tasa descuento 8%..20% step 1% (NPV base scenario)
discount_rates = [i/100 for i in range(8, 21)]
sens_disc_rows = []
base_params = scenarios["Base"]
for r in discount_rates:
    compra_cf = scenario_results["Base"]["compra_cf"]
    renta_cf = scenario_results["Base"]["renta_cf"]
    leasing_cf = scenario_results["Base"]["leasing_cf"]
    sens_disc_rows.append({
        "Tasa": r,
        "NPV_Compra": npv(r, compra_cf),
        "NPV_Renta": npv(r, renta_cf),
        "NPV_Leasing": npv(r, leasing_cf)
    })
df_sens_disc = pd.DataFrame(sens_disc_rows)

# Sensitivity 2: renta mensual por unidad 40000..60000 step 2000
renta_values = list(range(40000, 60001, 2000))
sens_renta_rows = []
for rv in renta_values:
    _, renta_cf_var, _ = generate_flows_for_params(
        base_params["costo_unit_usd"],
        base_params["tipo_cambio"],
        rv,
        base_params["mantenimiento_pct"],
        base_params["battery_pct"],
        lease_monthly_fleet_local = scenario_results["Base"]["lease_monthly_fleet_est"],
        residual_total_local = residual_total
    )
    sens_renta_rows.append({
        "RentaUnit": rv,
        "NPV_Renta": npv(base_params["tasa_descuento"], renta_cf_var)
    })
df_sens_renta = pd.DataFrame(sens_renta_rows)

# Sensitivity 3: costo equipo en USD 240k..320k step 8000
cost_values = list(range(240000, 320001, 8000))
sens_cost_rows = []
for cu in cost_values:
    compra_cf_var, _, _ = generate_flows_for_params(
        cu,
        base_params["tipo_cambio"],
        base_params["renta_mensual_unit"],
        base_params["mantenimiento_pct"],
        base_params["battery_pct"],
        lease_monthly_fleet_local = scenario_results["Base"]["lease_monthly_fleet_est"],
        residual_total_local = cu * base_params["tipo_cambio"] * units * 0.10
    )
    sens_cost_rows.append({
        "CostoUnitUSD": cu,
        "NPV_Compra": npv(base_params["tasa_descuento"], compra_cf_var)
    })
df_sens_cost = pd.DataFrame(sens_cost_rows)

# -----------------------------
# 6) CREAR EXCEL CON XlsxWriter
# -----------------------------
xlsx_filename = "Modelo_Financiero_Completo.xlsx"
writer = pd.ExcelWriter(xlsx_filename, engine='xlsxwriter')
workbook = writer.book
fmt_bold = workbook.add_format({"bold": True})
fmt_money = workbook.add_format({"num_format": '#,##0',})
fmt_money_dec = workbook.add_format({"num_format": '#,##0.00',})
fmt_pct = workbook.add_format({"num_format": '0.00%'})

# Supuestos
df_supuestos.to_excel(writer, sheet_name="Supuestos", index=False, startrow=0)
sheet_s = writer.sheets["Supuestos"]
sheet_s.set_column("A:A", 48)
sheet_s.set_column("B:B", 20)

# Flujos (todas las alternativas y escenarios)
df_flows.to_excel(writer, sheet_name="Flujos", index=False, startrow=0)
sheet_f = writer.sheets["Flujos"]
sheet_f.set_column("A:A", 12)
sheet_f.set_column("B:B", 6)
sheet_f.set_column("C:E", 18)
# Format flows as money
sheet_f.set_column("C:E", 18, fmt_money)

# Graf (acumulados)
df_graf.to_excel(writer, sheet_name="Grafico", index=False, startrow=0)
sheet_g = writer.sheets["Grafico"]
sheet_g.set_column("A:A", 12)
sheet_g.set_column("B:D", 18, fmt_money)

# Resumen (NPV & IRR)
df_summary.to_excel(writer, sheet_name="Resumen", index=False, startrow=0)
sheet_r = writer.sheets["Resumen"]
sheet_r.set_column("A:A", 12)
sheet_r.set_column("B:D", 18, fmt_money)
sheet_r.set_column("E:G", 12, fmt_money_dec)
sheet_r.set_column("H:I", 18, fmt_money)

# Sensibilidad
df_sens_disc.to_excel(writer, sheet_name="Sensibilidad", index=False, startrow=0)
df_sens_renta.to_excel(writer, sheet_name="Sensibilidad", index=False, startrow=len(df_sens_disc)+3, header=True)
df_sens_cost.to_excel(writer, sheet_name="Sensibilidad", index=False, startrow=len(df_sens_disc)+len(df_sens_renta)+7, header=True)
sheet_sens = writer.sheets["Sensibilidad"]
sheet_sens.set_column(0, 6, 18, fmt_money)

# Dashboard
sheet_dash = workbook.add_worksheet("Dashboard")
sheet_dash.set_column("A:A", 36)
sheet_dash.set_column("B:B", 20)

# Write KPI table header
sheet_dash.write("A1", "Dashboard - KPIs (por Escenario)", fmt_bold)
sheet_dash.write("A3", "Seleccionar Escenario", fmt_bold)
# data validation dropdown listing scenarios
sc_list = list(scenarios.keys())
# write scenario names in a hidden area (we'll put in sheet below row 40)
for i, s in enumerate(sc_list):
    sheet_dash.write(40 + i, 0, s)
# data validation cell
sheet_dash.data_validation('B3', {'validate': 'list', 'source': [s for s in sc_list]})
# headers
sheet_dash.write("A5", "Escenario", fmt_bold)
sheet_dash.write("B5", "NPV Compra (MXN)", fmt_bold)
sheet_dash.write("C5", "NPV Renta (MXN)", fmt_bold)
sheet_dash.write("D5", "NPV Leasing (MXN)", fmt_bold)
sheet_dash.write("E5", "IRR Compra", fmt_bold)
sheet_dash.write("F5", "IRR Renta", fmt_bold)
sheet_dash.write("G5", "IRR Leasing", fmt_bold)
sheet_dash.write("H5", "Lease Mensual unidad (MXN)", fmt_bold)

# write summary table into Dashboard area (we will also show chart using this table)
for i, row in df_summary.iterrows():
    sheet_dash.write(i+6, 0, row["Escenario"])
    sheet_dash.write_number(i+6, 1, row["NPV_Compra"], fmt_money)
    sheet_dash.write_number(i+6, 2, row["NPV_Renta"], fmt_money)
    sheet_dash.write_number(i+6, 3, row["NPV_Leasing"], fmt_money)
    # IRRs may be None
    irr_c = row["IRR_Compra"]
    irr_r = row["IRR_Renta"]
    irr_l = row["IRR_Leasing"]
    sheet_dash.write_number(i+6, 4, irr_c if irr_c is not None else 0, fmt_money_dec)
    sheet_dash.write_number(i+6, 5, irr_r if irr_r is not None else 0, fmt_money_dec)
    sheet_dash.write_number(i+6, 6, irr_l if irr_l is not None else 0, fmt_money_dec)
    sheet_dash.write_number(i+6, 7, row["Lease_Monthly_Unit"], fmt_money)

# Add a chart on dashboard: bar chart of NPVs by scenario (for each alternative)
chart_npv = workbook.add_chart({'type': 'column'})
# series Compra (located in Dashboard sheet)
n = len(df_summary)
chart_npv.add_series({'name': 'NPV Compra', 'categories': f'=Dashboard!$A$7:$A${6+n}', 'values': f'=Dashboard!$B$7:$B${6+n}'})
chart_npv.add_series({'name': 'NPV Renta', 'categories': f'=Dashboard!$A$7:$A${6+n}', 'values': f'=Dashboard!$C$7:$C${6+n}'})
chart_npv.add_series({'name': 'NPV Leasing', 'categories': f'=Dashboard!$A$7:$A${6+n}', 'values': f'=Dashboard!$D$7:$D${6+n}'})
chart_npv.set_title({'name': 'NPV por Escenario y Alternativa (MXN)'})
chart_npv.set_x_axis({'name': 'Escenario'})
chart_npv.set_y_axis({'name': 'NPV (MXN)'})
chart_npv.set_style(10)
sheet_dash.insert_chart('A12', chart_npv, {'x_scale': 1.3, 'y_scale': 1.1})

# Add cost accumulation chart (we'll use Grafico sheet series for Base scenario as main)
chart_cost = workbook.add_chart({'type': 'line'})
# find rows in df_graf for Base scenario
base_graf = df_graf[df_graf["Escenario"] == "Base"]
# write a small table on Dashboard for plotting base accum series
sheet_dash.write("J3", "Año", fmt_bold)
for i, year in enumerate(base_graf["Año"].tolist()):
    sheet_dash.write_number(3 + i + 1, 9, int(year))
sheet_dash.write("K3", "Compra Acum", fmt_bold)
sheet_dash.write("L3", "Renta Acum", fmt_bold)
sheet_dash.write("M3", "Leasing Acum", fmt_bold)
for i, row in enumerate(base_graf.itertuples()):
    sheet_dash.write_number(4 + i, 10, row.Compra_Acumulado, fmt_money)
    sheet_dash.write_number(4 + i, 11, row.Renta_Acumulado, fmt_money)
    sheet_dash.write_number(4 + i, 12, row.Leasing_Acumulado, fmt_money)

n_points = len(base_graf)
chart_cost.add_series({'name': 'Compra Acum', 'categories': f'=Dashboard!$J$4:$J${3+n_points}', 'values': f'=Dashboard!$K$4:$K${3+n_points}'})
chart_cost.add_series({'name': 'Renta Acum', 'categories': f'=Dashboard!$J$4:$J${3+n_points}', 'values': f'=Dashboard!$L$4:$L${3+n_points}'})
chart_cost.add_series({'name': 'Leasing Acum', 'categories': f'=Dashboard!$J$4:$J${3+n_points}', 'values': f'=Dashboard!$M$4:$M${3+n_points}'})
chart_cost.set_title({'name': 'Costo acumulado - Escenario Base (MXN)'})
chart_cost.set_x_axis({'name': 'Año'})
chart_cost.set_y_axis({'name': 'Costo acumulado (MXN)'})
sheet_dash.insert_chart('A30', chart_cost, {'x_scale': 1.5, 'y_scale': 1.2})

# Add text recommendation box (simple heuristic: choose alternative with highest (least negative) NPV per scenario)
sheet_dash.write("A55", "Recomendación automática (heurística):", fmt_bold)
for i, row in df_summary.iterrows():
    arr = {"Compra": row["NPV_Compra"], "Renta": row["NPV_Renta"], "Leasing": row["NPV_Leasing"]}
    # higher NPV (less negative) is better
    best = max(arr.items(), key=lambda x: x[1])[0]
    sheet_dash.write(i+56, 0, f"{row['Escenario']}: Mejor alternativa -> {best}")

# Create additional sheet with CSV-friendly tables for export
df_flows.to_excel(writer, sheet_name="Export_Flujos", index=False)

# Save workbook
writer.save()

# -----------------------------
# 7) GENERAR CSVs Y ZIP
# -----------------------------
zip_filename = "FinanIA_Analisis_Montacargas_Completo.zip"
with zipfile.ZipFile(zip_filename, 'w', zipfile.ZIP_DEFLATED) as zf:
    # CSVs
    zf.writestr("1_Supuestos.csv", df_supuestos.to_csv(index=False).encode('utf-8'))
    zf.writestr("2_Flujos.csv", df_flows.to_csv(index=False).encode('utf-8'))
    zf.writestr("3_Resumen.csv", df_summary.to_csv(index=False).encode('utf-8'))
    zf.writestr("4_Grafico_Acumulado.csv", df_graf.to_csv(index=False).encode('utf-8'))
    zf.writestr("5_Sensibilidad_Disc.csv", df_sens_disc.to_csv(index=False).encode('utf-8'))
    zf.writestr("6_Sensibilidad_Renta.csv", df_sens_renta.to_csv(index=False).encode('utf-8'))
    zf.writestr("7_Sensibilidad_Costo.csv", df_sens_cost.to_csv(index=False).encode('utf-8'))
    # add the xlsx
    with open(xlsx_filename, 'rb') as f:
        zf.writestr(xlsx_filename, f.read())

print("Generado exitosamente:")
print(" -", xlsx_filename)
print(" -", zip_filename)
print("\nAbre el Excel y revisa la hoja 'Dashboard' para KPIs y gráficos. Las hojas 'Resumen' y 'Sensibilidad' contienen tablas y fórmulas adicionales.")
