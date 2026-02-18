import pandas as pd
import numpy as np
import requests
import os


def run_analysis(excel_file_path, start_date, end_date):

    summary_rows = []
    output_folder = "outputs"
    os.makedirs(output_folder, exist_ok=True)

    file_name = os.path.join(output_folder, "MF_Combined_NAV_YOY_Screening.xlsx")

    scheme_df = pd.read_excel(excel_file_path, dtype=str)
    scheme_df.columns = scheme_df.columns.str.strip()

    if 'Scheme Code' in scheme_df.columns:
        scheme_df.rename(columns={'Scheme Code': 'SchemeCode'}, inplace=True)

    required_cols = {"SchemeCode", "AMC"}
    if not required_cols.issubset(scheme_df.columns):
        return None, "Excel must contain columns: SchemeCode, AMC"

    scheme_list = scheme_df.to_dict("records")

    with pd.ExcelWriter(file_name, engine="xlsxwriter") as writer:

        for row in scheme_list:
            scheme_code = row["SchemeCode"]
            amc = row["AMC"]

            url = f"https://api.mfapi.in/mf/{scheme_code}"
            response = requests.get(url)

            if response.status_code != 200:
                continue

            data = response.json()
            if 'data' not in data or len(data['data']) < 400:
                continue

            df = pd.DataFrame(data['data'])
            df['date'] = pd.to_datetime(df['date'], dayfirst=True)
            df['nav'] = df['nav'].astype(float)

            start_dt = pd.to_datetime(start_date)
            end_dt = pd.to_datetime(end_date)

            df_filtered = df[(df['date'] >= start_dt) & (df['date'] <= end_dt)]

            if len(df_filtered) < 400:
                continue

            monthly = (
                df_filtered
                .set_index('date')
                .resample('MS')
                .first()
                .reset_index()
            )

            yoy = pd.DataFrame({
                'Nav from': monthly['nav'],
                'Nav to': monthly['nav'].shift(-12)
            }).dropna()

            yoy['growth'] = ((yoy['Nav to'] / yoy['Nav from']) - 1) * 100

            total = len(yoy)
            positive_pct = round((yoy['growth'] > 0).mean() * 100, 2)
            negative_pct = round((yoy['growth'] <= 0).mean() * 100, 2)
            worst = round(yoy['growth'].min(), 2)
            best = round(yoy['growth'].max(), 2)
            avg = round(yoy['growth'].mean(), 2)
            volatility = round(yoy['growth'].std(), 2)

            verdict = (
                "PASS" if positive_pct >= 70 and worst > -25
                else "WATCH" if positive_pct >= 55
                else "REJECT"
            )

            summary_rows.append({
                "Scheme Code": scheme_code,
                "AMC": amc,
                "Positive YoY %": positive_pct,
                "Negative YoY %": negative_pct,
                "Worst YoY %": worst,
                "Best YoY %": best,
                "Average YoY %": avg,
                "Volatility": volatility,
                "Verdict": verdict
            })

        summary_df = pd.DataFrame(summary_rows)
        summary_df.to_excel(writer, sheet_name="SCREENING_SUMMARY", index=False)

    return summary_df, file_name
