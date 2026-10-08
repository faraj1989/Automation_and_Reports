import streamlit as st

import cached
import monitor

rows = cached.freshness()
stale = [r for r in rows if r["stale"]]
if stale:
    st.warning(f"{len(stale)} of {len(rows)} sources are older than expected.", icon=":material/warning:")
else:
    st.success(f"All {len(rows)} sources are fresh.", icon=":material/check_circle:")

st.dataframe(
    [{"Status": "Stale" if r["stale"] else "OK", "Source": r["source"], "Updated": r["newest"],
      "Age": r["age"], "Expected within (h)": r["max_age_h"],
      "Size": monitor.human_bytes(r["size"]) if r.get("size") else "", "Newest file": r["file"]}
     for r in sorted(rows, key=lambda r: not r["stale"])],
    hide_index=True,
    column_config={"Updated": st.column_config.DatetimeColumn(format="YYYY-MM-DD HH:mm")},
)
st.caption("Alarm exports should change every few minutes, daily outputs within ~30 h, hourly cell files "
           "within ~7 h, weekly/monthly reports within their period.")

st.subheader("Disks")
for anchor, du in cached.system()["drives"].items():
    st.progress(du["used"] / du["total"],
                text=f"{anchor} - {monitor.human_bytes(du['free'])} free of {monitor.human_bytes(du['total'])}")
