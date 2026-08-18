#!/usr/bin/env python3
"""
generate_data.py
Reads data/accounts.xlsx (SFDC export from Quest) and outputs docs/data.json.
Run locally or via GitHub Actions.
"""

import json
import re
import sys
from datetime import date
from pathlib import Path

import pandas as pd

ROOT = Path(__file__).parent.parent
INPUT  = ROOT / "data" / "accounts.xlsx"
OUTPUT = ROOT / "docs" / "data.json"

# ── Account name normalization ────────────────────────────────────────────────
NORM_MAP = {
    'AIG':                          lambda n: 'AIG' in n,
    'UNISYS CORP':                  lambda n: 'UNISYS' in n,
    'S&P GLOBAL':                   lambda n: 'S&P GLOBAL' in n,
    'SFBCIC':                       lambda n: 'SOUTHERN FARM BUREAU CASUALTY' in n,
    'SFBLIC':                       lambda n: 'SOUTHERN FARM BUREAU LIFE' in n,
}

def normalize(name: str) -> str:
    if not isinstance(name, str):
        return str(name)
    upper = name.upper()
    for display, test in NORM_MAP.items():
        if test(upper):
            return display
    return name


def main():
    if not INPUT.exists():
        print(f"ERROR: {INPUT} not found.", file=sys.stderr)
        sys.exit(1)

    print(f"Reading {INPUT} ...")
    raw = pd.read_excel(INPUT, header=None)

    # Find header row (contains "SFDC Account ID")
    header_row = None
    for i, row in raw.iterrows():
        vals = [str(v) for v in row if str(v) != 'nan']
        if 'SFDC Account ID' in vals:
            header_row = i
            break

    if header_row is None:
        print("ERROR: Could not find header row with 'SFDC Account ID'.", file=sys.stderr)
        sys.exit(1)

    df = pd.read_excel(INPUT, header=header_row)

    # Rename columns
    col_map = {
        'SFDC Account ID':                                              'sfdc_id',
        'Account Name':                                                 'account_name',
        'Account : Account Owner : Business Unit':                      'owner_bu',
        'Product Description':                                          'product',
        'Quantity':                                                     'quantity',
        'Maint. End Date':                                              'maint_end',
        'Asset Number':                                                 'asset_number',
        'Asset : Contact : Email':                                      'contact_email',
        'Quest Product Family':                                         'product_family',
        'Business Unit':                                                'business_unit',
        'Account : Account Owner : Enter Oppty Amount':                 'oppty_amount',
        'Order Product: Total Amount':                                  'total_amount',
        'Partner Status':                                               'partner_status',
        'Asset : Order Product : Order : Opportunity : Quest Primary Partner': 'partner',
        'Account : Account Owner : Username':                           'owner_username',
    }
    df = df.rename(columns={k: v for k, v in col_map.items() if k in df.columns})

    wanted = list(col_map.values())
    df = df[[c for c in wanted if c in df.columns]]

    # Only real SFDC IDs
    df = df[df['sfdc_id'].astype(str).str.match(r'^Act-\d+', na=False)]
    df['account_display'] = df['account_name'].apply(normalize)
    df['maint_end_dt']    = pd.to_datetime(df['maint_end'], errors='coerce')

    today = pd.Timestamp(date.today())

    # ── Account-level summary ────────────────────────────────────────────────
    acct = df.groupby('account_display').agg(
        total_revenue=('total_amount',     'sum'),
        assets=       ('asset_number',     'count'),
        next_renewal= ('maint_end_dt',     'min'),
        products=     ('product_family',   lambda x: sorted(list(x.dropna().unique()))),
        partner=      ('partner',          lambda x: x.dropna().iloc[0] if x.dropna().any() else 'Direct'),
        sfdc_id=      ('sfdc_id',          'first'),
        contact=      ('contact_email',    lambda x: next((v for v in x if pd.notna(v) and v), '')),
    ).reset_index()

    acct['days_to_renewal'] = (acct['next_renewal'] - today).dt.days
    acct['next_renewal_str'] = acct['next_renewal'].dt.strftime('%m/%d/%Y')
    acct['renewal_risk'] = acct['days_to_renewal'].apply(
        lambda d: 'critical' if d < 60 else ('warning' if d < 120 else 'ok'))

    acct_list = []
    for _, row in acct.iterrows():
        acct_list.append({
            'name':           row['account_display'],
            'sfdc_id':        row['sfdc_id'],
            'total_revenue':  round(float(row['total_revenue']), 2),
            'assets':         int(row['assets']),
            'next_renewal':   row['next_renewal_str'],
            'days_to_renewal': int(row['days_to_renewal']) if pd.notna(row['days_to_renewal']) else None,
            'renewal_risk':   row['renewal_risk'],
            'products':       row['products'],
            'partner':        str(row['partner']) if pd.notna(row['partner']) else 'Direct',
            'contact':        str(row['contact'])  if pd.notna(row['contact'])  else '',
        })
    acct_list.sort(key=lambda x: x['total_revenue'], reverse=True)

    # ── Full asset detail per account ─────────────────────────────────────────
    all_assets: dict = {}
    for sfdc_id, group in df.groupby('sfdc_id'):
        acct_display = group['account_display'].iloc[0]
        assets = []
        for _, row in group.iterrows():
            assets.append({
                'product': str(row['product'])       if pd.notna(row['product'])       else '',
                'family':  str(row['product_family']) if pd.notna(row['product_family']) else '',
                'asset':   str(row['asset_number'])  if pd.notna(row['asset_number'])  else '',
                'qty':     int(row['quantity'])       if pd.notna(row['quantity'])       else 0,
                'maint':   str(row['maint_end'])      if pd.notna(row['maint_end'])      else '',
                'amount':  round(float(row['total_amount']), 2) if pd.notna(row['total_amount']) else 0.0,
                'contact': str(row['contact_email'])  if pd.notna(row['contact_email']) else '',
                'bu':      str(row['business_unit'])  if pd.notna(row['business_unit']) else '',
            })
        all_assets.setdefault(acct_display, []).extend(assets)

    # ── Aggregates ─────────────────────────────────────────────────────────────
    bu_breakdown = {
        k: round(float(v), 2)
        for k, v in df.groupby('business_unit')['total_amount'].sum().items()
    }
    pf_revenue = {
        k: round(float(v), 2)
        for k, v in df.groupby('product_family')['total_amount']
                      .sum().sort_values(ascending=False).head(10).items()
    }
    renewals = df[df['maint_end_dt'].notna()].copy()
    renewals['month'] = renewals['maint_end_dt'].dt.to_period('M')
    monthly_renewals = {
        str(k): round(float(v), 2)
        for k, v in renewals.groupby('month')['total_amount'].sum().items()
        if str(k) >= str(today.to_period('M'))
        and str(k) <= str((today + pd.DateOffset(months=12)).to_period('M'))
    }

    payload = {
        'accounts':         acct_list,
        'assets':           all_assets,
        'bu_breakdown':     bu_breakdown,
        'pf_revenue':       pf_revenue,
        'monthly_renewals': monthly_renewals,
        'total_revenue':    round(float(df['total_amount'].sum()), 2),
        'total_assets':     int(len(df)),
        'total_accounts':   int(df['account_display'].nunique()),
        'at_risk_count':    int((acct['renewal_risk'] == 'critical').sum()),
        'generated':        str(date.today()),
    }

    OUTPUT.parent.mkdir(parents=True, exist_ok=True)
    with open(OUTPUT, 'w') as f:
        json.dump(payload, f, separators=(',', ':'))

    print(f"✓ Wrote {OUTPUT} ({OUTPUT.stat().st_size:,} bytes)")
    print(f"  Accounts: {payload['total_accounts']}, Assets: {payload['total_assets']}, "
          f"At-risk: {payload['at_risk_count']}, Generated: {payload['generated']}")


if __name__ == '__main__':
    main()
