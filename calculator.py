# -*- coding: utf-8 -*-
"""Excel processing logic used by the Flask web app.

Cloud-friendly notes:
- No code runs on import.
- The entry point is process_excel(input_path, output_path).
"""

from __future__ import annotations

import os
import uuid
import shutil
import numpy as np
import pandas as pd

from openpyxl import load_workbook
from openpyxl.utils import get_column_letter


RIGA_TZ = "Europe/Riga"
HOURLY_RANGE = pd.date_range("00:00", "23:00", freq="1h").strftime("%H:%M").tolist()
TOTAL_LABEL = "TOTAL "  # trailing space is intentional to avoid collisions


def _parse_dt(series: pd.Series) -> pd.Series:
    """Parse datetime from 'dd.mm.yyyy HH:MM' or any parseable format."""
    # Use explicit format parsing with dayfirst=True for DD.MM.YYYY
    dt = pd.to_datetime(series, format="%d.%m.%Y %H:%M", errors="coerce", dayfirst=True)
    if dt.isna().all():
        dt = pd.to_datetime(series, dayfirst=True, errors="coerce")
    return dt


def _api_window(min_local: pd.Timestamp, max_local: pd.Timestamp):
    """Build a safe time window for Elering API in UTC ISO timestamps."""
    start_local = (min_local.normalize() - pd.Timedelta(days=1)).tz_localize(RIGA_TZ)
    end_local = (
        max_local.normalize() + pd.Timedelta(days=1) - pd.Timedelta(milliseconds=1)
    ).tz_localize(RIGA_TZ)

    start_utc = start_local.tz_convert("UTC").strftime("%Y-%m-%dT%H:%M:%S.%f")[:-3] + "Z"
    end_utc = end_local.tz_convert("UTC").strftime("%Y-%m-%dT%H:%M:%S.%f")[:-3] + "Z"
    return start_utc, end_utc


def _read_prices(start_utc: str, end_utc: str) -> pd.DataFrame:
    """Download price CSV from Elering and normalize columns."""
    url = (
        "https://dashboard.elering.ee/api/nps/price/csv?start="
        + start_utc.replace(":", "%3A")
        + "&end="
        + end_utc.replace(":", "%3A")
        + "&fields=lv"
    )
    print(url)

    df = pd.read_csv(
        url,
        encoding="latin-1",
        quotechar='"',
        delimiter=";",
        decimal=",",
    )

    df["dt_local"] = pd.to_datetime(
        df["Kuupäev (Eesti aeg)"],
        format="%d.%m.%Y %H:%M",
        errors="coerce",
    )

    # Pick Latvia column robustly (accent/no-accent)
    price_col = None
    for cand in ("NPS Läti", "NPS Lati", "NPS Latvia", "LV", "LV (EUR)"):
        if cand in df.columns:
            price_col = cand
            break
    if price_col is None:
        # Fallback: first column containing 'Lati' or 'Läti'
        for c in df.columns:
            cs = str(c)
            if "Lati" in cs or "Läti" in cs:
                price_col = c
                break
    if price_col is None:
        raise ValueError(f"Could not find Latvia price column in CSV. Columns: {list(df.columns)}")

    df = df.rename(columns={price_col: "price_eur_mwh"})[["dt_local", "price_eur_mwh"]].dropna(
        subset=["dt_local"]
    ).copy()

    df["price_eur_mwh"] = pd.to_numeric(df["price_eur_mwh"], errors="coerce")
    return df


def _detect_resolution(price_df: pd.DataFrame) -> str:
    """Detect if series is hourly (60min) or quarter-hour (15min).

    Robust: uses minimal non-zero step across the whole series.
    """
    price_df = price_df.sort_values("dt_local")
    if len(price_df) < 2:
        return "60min"

    dt = pd.to_datetime(price_df["dt_local"], errors="coerce").dropna()
    if len(dt) < 2:
        return "60min"

    diffs = dt.diff().dropna()
    diffs = diffs[diffs > pd.Timedelta(seconds=1)]
    if diffs.empty:
        return "60min"

    min_s = diffs.min().total_seconds()
    return "15min" if min_s < 1800 else "60min"


def _detect_series_resolution(dt_series: pd.Series) -> str:
    """Detect resolution of datetime series: '15min' or '60min'."""
    s = pd.to_datetime(dt_series, errors="coerce").dropna().sort_values().drop_duplicates()
    if len(s) < 2:
        return "60min"
    diffs = s.diff().dropna()
    diffs = diffs[diffs > pd.Timedelta(seconds=1)]
    if diffs.empty:
        return "60min"
    min_s = diffs.min().total_seconds()
    return "15min" if min_s < 1800 else "60min"


def _align(consumption: pd.DataFrame, prices: pd.DataFrame, price_resolution: str) -> pd.DataFrame:
    """Align consumption (sheet 'Dati') with prices.

    Key points:
    - If consumption is hourly but prices are 15-min, use HOURLY mean price for the hour.
      This matches Nord Pool 1 Hour price indices and avoids picking only HH:00 quarter price.
    - If consumption is 15-min and prices are hourly, merge by hour (price replicated to 15-min rows).
    """
    left = consumption.copy()
    left["dt_local"] = _parse_dt(left["Datums"])
    left = left.dropna(subset=["dt_local"])

    cons_resolution = _detect_series_resolution(left["dt_local"])

    price = prices.copy()
    price["dt_local"] = pd.to_datetime(price["dt_local"], errors="coerce")
    price = price.dropna(subset=["dt_local"]).copy()

    # Decide merge key and normalize price to match it
    if cons_resolution == "60min" and price_resolution == "15min":
        # Hourly consumption: use hourly mean price
        price["dt_key"] = price["dt_local"].dt.floor("h")
        price = price.groupby("dt_key", as_index=False)["price_eur_mwh"].mean()
        left["dt_key"] = left["dt_local"].dt.floor("h")
    elif price_resolution == "15min" and cons_resolution == "15min":
        left["dt_key"] = left["dt_local"].dt.floor("15min")
        price["dt_key"] = price["dt_local"].dt.floor("15min")
    else:
        # Default: hourly key
        left["dt_key"] = left["dt_local"].dt.floor("h")
        price["dt_key"] = price["dt_local"].dt.floor("h")
        # If multiple prices per hour (shouldn't happen here), take mean
        price = price.groupby("dt_key", as_index=False)["price_eur_mwh"].mean()

    merged = pd.merge(
        left,
        price[["dt_key", "price_eur_mwh"]],
        on="dt_key",
        how="left",
    ).drop(columns=["dt_key"])

    merged = merged.rename(columns={"price_eur_mwh": "Prices"})

    # Shift prices forward by 15 minutes to align with consumption timing
    # (consumption is recorded at HH:MM, prices apply at HH:MM+15min)
    merged["Prices"] = merged["Prices"].shift(1)

    return merged


def _blank_like(df: pd.DataFrame) -> dict:
    """Return a dict with blank values matching DataFrame dtypes."""
    return {
        col: (np.nan if pd.api.types.is_numeric_dtype(df[col]) else "")
        for col in df.columns
    }


def _by_days_hourly(df: pd.DataFrame, value_col: str, agg: str = "sum") -> pd.DataFrame:
    """Build ByDays-style block aggregated by hours (00:00–23:00)."""
    df = df.copy()
    df["dt_floor"] = pd.to_datetime(df["dt_local"]).dt.floor("h")
    df["date"] = df["dt_floor"].dt.date
    df["time"] = df["dt_floor"].dt.strftime("%H:%M")

    table = df.pivot_table(
        index="time",
        columns="date",
        values=value_col,
        aggfunc=agg,
    ).reindex(HOURLY_RANGE)

    if table is None:
        table = pd.DataFrame(index=HOURLY_RANGE)

    table.columns = [pd.to_datetime(c).strftime("%d.%m.%Y") for c in table.columns]
    table = table.reset_index().rename(columns={"time": "Time"})

    # Remove the next month's 00:00 entry (e.g., 01.09 when period is 01.08-31.08)
    cols_to_drop = []
    for col in table.columns:
        if col.startswith("01.") and col.endswith("2026"):
            # Check if this is the next month (September = 09)
            try:
                if ".09.2026" in col:
                    cols_to_drop.append(col)
            except:
                pass

    if cols_to_drop:
        table = table.drop(columns=cols_to_drop)

    # Empty row for visual separation
    table.loc[len(table)] = _blank_like(table)
    return table


def _add_total(block: pd.DataFrame) -> pd.DataFrame:
    """Add TOTAL row for a ByDays block and one more empty row after it."""
    if block is None or block.empty:
        return block

    sum_row = block.drop(columns=["Time"]).sum(axis=0, numeric_only=True)
    total_row = pd.DataFrame(
        [[TOTAL_LABEL] + sum_row.tolist()],
        columns=["Time"] + list(sum_row.index),
    )

    combined = pd.concat([block.iloc[:-1], total_row, block.iloc[-1:]], ignore_index=True)

    last_total_idx = int(combined.index[combined["Time"] == TOTAL_LABEL][-1])
    top = combined.iloc[: last_total_idx + 1].copy()
    bottom = combined.iloc[last_total_idx + 1 :].copy()
    top.loc[len(top)] = _blank_like(combined)

    return pd.concat([top, bottom], ignore_index=True)


def process_excel(input_path: str, output_path: str) -> None:
    """Process a single uploaded Excel file and write the result to output_path."""

    if not os.path.isfile(input_path):
        raise FileNotFoundError(f"Input file not found: {input_path}")

    os.makedirs(os.path.dirname(output_path) or ".", exist_ok=True)

    # Work on a temporary copy to keep the original upload intact.
    work_dir = os.path.dirname(output_path) or os.path.dirname(input_path) or "."
    work_path = os.path.join(work_dir, f"work_{uuid.uuid4().hex}.xlsx")
    shutil.copy2(input_path, work_path)

    try:
        df = pd.read_excel(work_path, sheet_name="Dati", skipfooter=1, engine=("xlrd" if str(work_path).lower().endswith(".xls") else "openpyxl"))

        obj_name = "UNKNOWN"
        if "Objekta adrese" in df.columns:
            try:
                obj_name = df["Objekta adrese"].astype(str).str.split().str[-1].unique()[0]
            except Exception:
                obj_name = "UNKNOWN"

        print("Object:", obj_name)

        dt = _parse_dt(df["Datums"])
        df["dt_local"] = dt
        min_date = dt.min().normalize()
        max_date = dt.max().normalize()

        print("Start (local) =", min_date.date(), "End (local) =", max_date.date())

        # Define end_of_month once to filter out next month's data
        # Use min_date to get the correct month (in case max_date extends into next month)
        end_of_month = min_date.replace(day=1) + pd.DateOffset(months=1) - pd.DateOffset(days=1)
        # Extend to include 23:59:59 of the last day
        end_of_month = end_of_month.replace(hour=23, minute=59, second=59)

        start_utc, end_utc = _api_window(min_date, max_date)
        price_df = _read_prices(start_utc, end_utc)

        resolution = _detect_resolution(price_df)
        print("Detected price resolution:", resolution)

        merged = _align(df, price_df, resolution)

        # Calculate energy (A+) and related sums
        if "Patērētā elektroenerģija (A+) (kWh)" in merged.columns:
            merged["InCalculatedSum"] = (
                merged["Patērētā elektroenerģija (A+) (kWh)"] * merged["Prices"]
            ) / 1000.0

        # Calculate energy (A-) and related sums
        if "Tīklā nodotā enerģija (A-) (kWh)" in merged.columns:
            merged["OutCalculatedSum"] = (
                merged["Tīklā nodotā enerģija (A-) (kWh)"].fillna(0) * merged["Prices"]
            ) / 1000.0

        # Now filter valid records and calculate TOTAL row
        valid_records = merged[merged["dt_local"] <= end_of_month]

        if "Patērētā elektroenerģija (A+) (kWh)" in merged.columns:
            merged.at[TOTAL_LABEL, "Patērētā elektroenerģija (A+) (kWh)"] = round(
                valid_records["Patērētā elektroenerģija (A+) (kWh)"].sum(), 4
            )
            merged.at[TOTAL_LABEL, "InCalculatedSum"] = round(valid_records["InCalculatedSum"].sum(), 4)

        if "Tīklā nodotā enerģija (A-) (kWh)" in merged.columns:
            merged.at[TOTAL_LABEL, "Tīklā nodotā enerģija (A-) (kWh)"] = round(
                valid_records["Tīklā nodotā enerģija (A-) (kWh)"].sum(), 5
            )
            merged.at[TOTAL_LABEL, "OutCalculatedSum"] = round(valid_records["OutCalculatedSum"].sum(), 4)

        merged["Prices"] = pd.to_numeric(merged["Prices"], errors="coerce")

        # "Generated (EUR)" value that app.py shows in the History table.
        # IMPORTANT: we add a TOTAL row into the dataframe using merged.at[TOTAL_LABEL, ...].
        # If we blindly sum the whole column after that, the TOTAL row is counted too -> numbers become x2.
        # So we read the value from the TOTAL row (preferred), or sum excluding TOTAL_LABEL.
        total_generated = 0.0

        if "OutCalculatedSum" in merged.columns:
            if TOTAL_LABEL in merged.index:
                v = pd.to_numeric(merged.at[TOTAL_LABEL, "OutCalculatedSum"], errors="coerce")
                total_generated = float(v) if pd.notna(v) else 0.0
            else:
                total_generated = float(
                    pd.to_numeric(
                        merged.loc[merged.index != TOTAL_LABEL, "OutCalculatedSum"], errors="coerce"
                    ).sum()
                )
        elif "InCalculatedSum" in merged.columns:
            if TOTAL_LABEL in merged.index:
                v = pd.to_numeric(merged.at[TOTAL_LABEL, "InCalculatedSum"], errors="coerce")
                total_generated = float(v) if pd.notna(v) else 0.0
            else:
                total_generated = float(
                    pd.to_numeric(
                        merged.loc[merged.index != TOTAL_LABEL, "InCalculatedSum"], errors="coerce"
                    ).sum()
                )

        print(f"Total generated sum = {total_generated:.6f}")

        # ByDays (always hourly 00:00–23:00)
        prices_block = _by_days_hourly(merged[["dt_local", "Prices"]].dropna(), "Prices", agg="mean")

        cons_block = None
        if "Tīklā nodotā enerģija (A-) (kWh)" in merged.columns:
            # Filter out records from the next month (only for ByDays calculation)
            # Keep only records within the current month (using end_of_month defined above)
            filtered_data = merged[merged["dt_local"] <= end_of_month][["dt_local", "Tīklā nodotā enerģija (A-) (kWh)"]].rename(
                columns={"Tīklā nodotā enerģija (A-) (kWh)": "Aminus"}
            ).dropna()

            cons_block = _by_days_hourly(filtered_data, "Aminus")
            cons_block = _add_total(cons_block)

        calc_block = None
        if "OutCalculatedSum" in merged.columns:
            calc_block = _by_days_hourly(merged[["dt_local", "OutCalculatedSum"]], "OutCalculatedSum")
            calc_block = _add_total(calc_block)

        # Add row sum columns before concatenation (only for TOTAL rows)
        def add_sum_column(block, sum_col_name="Sum"):
            if block is None:
                return None
            date_cols = [c for c in block.columns if c != "Time"]
            # Calculate sums but only set them for TOTAL rows
            block[sum_col_name] = pd.NA
            for idx, row in block.iterrows():
                if row.get("Time") == TOTAL_LABEL:
                    row_sum = pd.to_numeric(row[date_cols], errors='coerce').sum()
                    block.loc[idx, sum_col_name] = row_sum
            return block

        prices_block = add_sum_column(prices_block)
        cons_block = add_sum_column(cons_block)
        calc_block = add_sum_column(calc_block)

        blocks = [prices_block]
        if cons_block is not None:
            blocks.append(cons_block)
        if calc_block is not None:
            blocks.append(calc_block)

        by_days = pd.concat(blocks, axis=0, ignore_index=True)

        by_cols = [c for c in by_days.columns if c not in ("Time", " ")]
        by_days[" "] = pd.NA
        by_days.loc[by_days["Time"] == TOTAL_LABEL, " "] = by_days.loc[
            by_days["Time"] == TOTAL_LABEL, by_cols
        ].sum(axis=1, numeric_only=True)

        # Remove the temporary total column before export
        by_days = by_days.drop(columns=[" "])

        # Write back to Excel
        wb = load_workbook(work_path)
        if "Total_Calculation" in wb.sheetnames:
            del wb["Total_Calculation"]
        if "ByDays" in wb.sheetnames:
            del wb["ByDays"]
        if "Dati" in wb.sheetnames:
            del wb["Dati"]
        wb.save(work_path)

        with pd.ExcelWriter(work_path, mode="a", engine="openpyxl") as writer:
            # Filter out records from next month, but keep TOTAL row
            merged_filtered = merged[(merged["dt_local"] <= end_of_month) | (merged.index == TOTAL_LABEL)].drop(columns=["dt_local"], errors="ignore")

            # Export Dati (same as Total_Calculation)
            merged_filtered.to_excel(writer, sheet_name="Dati", index=False)

            # Export Total_Calculation
            merged_filtered.to_excel(
                writer, sheet_name="Total_Calculation", index=False
            )
            by_days.to_excel(writer, sheet_name="ByDays", index=False)

        wb = load_workbook(work_path)

        # Apply optimal column widths for all sheets
        for sheet_name in wb.sheetnames:
            ws = wb[sheet_name]
            for column in ws.columns:
                max_length = 0
                column_letter = get_column_letter(column[0].column)

                for cell in column:
                    try:
                        if cell.value:
                            max_length = max(max_length, len(str(cell.value)))
                    except:
                        pass

                # Optimal width per sheet
                if sheet_name == "ByDays":
                    adjusted_width = min(max_length + 1, 12)
                elif sheet_name == "Total_Calculation":
                    adjusted_width = min(max_length + 1, 18)
                else:
                    adjusted_width = min(max_length + 2, 20)

                ws.column_dimensions[column_letter].width = adjusted_width

        wb.save(output_path)
        wb.close()

        print("Saved to:", output_path)

    finally:
        try:
            os.remove(work_path)
        except Exception:
            pass


if __name__ == "__main__":
    # Manual CLI run for debugging.
    import argparse

    parser = argparse.ArgumentParser()
    parser.add_argument("input", help="Path to input .xlsx")
    parser.add_argument("output", help="Path to output .xlsx")
    args = parser.parse_args()

    process_excel(args.input, args.output)
