"""
NAV Refresh — Streamlit front-end for refresh_nav.py

Two workflows, no local files, no PowerShell:
  1. Resolve fund names -> scheme codes (paste names, get ranked candidates)
  2. Refresh a workbook  (upload .xlsx, get a refreshed .xlsx back)

All data pulled is public (AMFI / mfapi.in NAV data) — nothing client-specific
is uploaded, stored, or displayed by this app.
"""

from __future__ import annotations

import io
import sys
import time

import openpyxl
import pandas as pd
import streamlit as st

import refresh_nav as nav

st.set_page_config(page_title="MF NAV Refresh", page_icon="📈", layout="centered")

st.title("📈 MF NAV & Returns — Refresh Tool")
st.caption(
    "Pulls NAV data from mfapi.in / AMFI (public sources only). "
    "Nothing you upload here is stored — the workbook is processed in "
    "memory and only exists in this browser tab."
)

mode = st.sidebar.radio(
    "What do you want to do?",
    ["Refresh a workbook", "Resolve fund names → scheme codes", "Source health check"],
)

# ---------------------------------------------------------------------------
# MODE 1: Refresh a workbook
# ---------------------------------------------------------------------------
if mode == "Refresh a workbook":
    st.subheader("Refresh a workbook")
    st.write(
        "Upload your `MF_NAV_Master.xlsx` (must already have the "
        "**Scheme Codes / NAV Anchors / NAVs / Returns / Refresh Log** tabs — "
        "use `refresh_nav.py --init` once to create a fresh one if you don't "
        "have it yet)."
    )

    uploaded = st.file_uploader("Workbook (.xlsx)", type=["xlsx"])
    allow_hist = st.checkbox(
        "Allow AMFI history fallback if mfapi.in fails for a fund (slower — "
        "pulls history in 90-day chunks)",
        value=False,
    )

    if uploaded is not None:
        if st.button("Refresh NAVs", type="primary"):
            try:
                wb = openpyxl.load_workbook(io.BytesIO(uploaded.getvalue()))
            except Exception as exc:
                st.error(f"Couldn't open that file as an .xlsx workbook: {exc}")
                st.stop()

            try:
                codes_preview = nav.read_codes(wb["Scheme Codes"])
            except KeyError:
                st.error(
                    "This workbook doesn't have a 'Scheme Codes' tab. "
                    "It doesn't look like a workbook built by refresh_nav.py."
                )
                st.stop()

            if not codes_preview:
                st.error("No scheme codes found in the 'Scheme Codes' tab.")
                st.stop()

            total = len(codes_preview)
            progress_bar = st.progress(0.0)
            status_area = st.empty()
            log_rows = []

            def on_progress(i, n, code, status, name):
                log_rows.append({"Code": code, "Fund": str(name)[:60], "Status": status})
                progress_bar.progress(i / n)
                status_area.write(f"{i}/{n} — `{code}` {name} → **{status}**")

            t0 = time.time()
            try:
                wb, ok, fail = nav.refresh_workbook(
                    wb, allow_amfi_history=allow_hist, on_progress=on_progress
                )
            except ValueError as exc:
                st.error(str(exc))
                st.stop()
            except Exception as exc:  # noqa: BLE001
                st.error(f"Refresh failed: {exc}")
                st.stop()

            elapsed = time.time() - t0
            progress_bar.progress(1.0)

            if fail == 0:
                st.success(f"Done in {elapsed:.0f}s — {ok} ok / {fail} failed.")
            else:
                st.warning(f"Done in {elapsed:.0f}s — {ok} ok / {fail} failed.")

            st.dataframe(pd.DataFrame(log_rows), use_container_width=True, hide_index=True)

            out = io.BytesIO()
            wb.save(out)
            out.seek(0)
            st.download_button(
                "⬇️ Download refreshed workbook",
                data=out,
                file_name=uploaded.name,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary",
            )

# ---------------------------------------------------------------------------
# MODE 2: Resolve names -> scheme codes
# ---------------------------------------------------------------------------
elif mode == "Resolve fund names → scheme codes":
    st.subheader("Resolve fund names → scheme codes")
    st.write(
        "Paste fund names (one per line). This only *searches* mfapi.in and "
        "ranks candidates — nothing gets written anywhere. Copy the correct "
        "code into your workbook's 'Scheme Codes' tab yourself."
    )

    names_text = st.text_area(
        "Fund names (one per line)",
        height=180,
        placeholder="Axis Small Cap Direct-G\nHDFC Flexi Cap Reg-G\n...",
    )

    if st.button("Resolve", type="primary"):
        names = [ln.strip() for ln in names_text.splitlines() if ln.strip()]
        if not names:
            st.warning("Paste at least one fund name first.")
            st.stop()

        with st.spinner(f"Searching mfapi.in for {len(names)} fund(s)..."):
            try:
                report = nav.resolve_names(names)
            except Exception as exc:  # noqa: BLE001
                st.error(f"Resolve failed: {exc}")
                st.stop()

        for row in report:
            st.markdown(f"**{row['requested']}**  \n*searched: \"{row['query']}\"*")
            if row["status"] == "no_results":
                st.warning("No matches — check spelling, or the fund may not be indexed.")
            elif row["status"] != "ok":
                st.error(f"Error: {row['status']}")
            else:
                cand_df = pd.DataFrame(row["candidates"])
                cand_df["score"] = (cand_df["score"] * 100).round(0).astype(int).astype(str) + "%"
                cand_df = cand_df.rename(
                    columns={"code": "Scheme Code", "name": "Fund Name", "score": "Match"}
                )
                st.dataframe(cand_df, use_container_width=True, hide_index=True)
                if row["candidates"][0]["score"] < 0.55:
                    st.caption(
                        "⚠️ No strong match — verify manually on mfapi.in or AMFI before using."
                    )
            st.divider()

# ---------------------------------------------------------------------------
# MODE 3: Health check
# ---------------------------------------------------------------------------
else:
    st.subheader("Source health check")
    st.write("Probes mfapi.in and both AMFI endpoints and reports what's currently reachable.")

    if st.button("Run health check", type="primary"):
        buf = io.StringIO()
        old_stdout = sys.stdout
        sys.stdout = buf
        try:
            problems = nav.healthcheck()
        finally:
            sys.stdout = old_stdout

        st.code(buf.getvalue(), language=None)
        if problems:
            st.error(f"{problems} source(s) failed.")
        else:
            st.success("All sources reachable.")
