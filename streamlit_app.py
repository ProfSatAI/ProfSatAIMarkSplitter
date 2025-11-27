import streamlit as st
import csv
import io
import random as ra
from typing import List, Tuple, Optional


# ------------------ CORE LOGIC ------------------ #

def get_marks_from_file(uploaded_file) -> List[int]:
    """
    Read total marks from uploaded CSV file.
    Assumes first column = total marks.
    Skips non-integer header rows automatically.
    """
    marks: List[int] = []
    # Wrap uploaded file (bytes) as text stream
    uploaded_file.seek(0)
    text_stream = io.TextIOWrapper(uploaded_file, encoding="utf-8")
    reader = csv.reader(text_stream)
    for row in reader:
        if not row:
            continue
        try:
            mark = int(row[0])
        except ValueError:
            # header or invalid cell
            continue
        marks.append(mark)
    return marks


def assessment_pattern(reg: int, ass: int, dep: Optional[int] = None) -> Tuple[List[int], List[int]]:
    """
    Returns (ms, co) lists for a given regulation and assessment.
    dep: 1 -> S&H, 2 -> Other (only for MODEL).
    NOTE: ass == 6 (Custom) is handled separately via custom_ms/custom_co.
    """
    ms: List[int] = []
    co: List[int] = []

    if reg in (13, 17):
        if ass == 1:
            ms = [2, 2, 2, 2, 2, 16, 16, 8]
            co = [1, 1, 1, 2, 2, 1, 2, 1]
        elif ass == 2:
            ms = [2, 2, 2, 2, 2, 16, 16, 8]
            co = [3, 3, 3, 4, 4, 3, 4, 3]
        elif ass == 3:
            if dep is None:
                raise ValueError("Department is required for MODEL exam.")
            if dep == 1:
                ms = [2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 16, 16, 16, 16, 16]
                co = [1, 1, 2, 2, 3, 3, 4, 4, 5, 5, 1, 2, 3, 4, 5]
            elif dep == 2:
                ms = [2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 13, 13, 13, 13, 13, 15]
                co = [1, 1, 2, 2, 3, 3, 4, 4, 5, 5, 1, 2, 3, 4, 5, 5]
            else:
                raise ValueError("Invalid department selected.")
        elif ass == 4:
            ms = [20, 20, 20, 20, 20]
            co = [1, 2, 3, 4, 5]
        elif ass == 5:
            ms = [20, 20, 20, 20, 20]
            co = [1, 2, 3, 4, 5]
        else:
            raise ValueError("Invalid assessment number for this regulation.")

    elif reg == 21:
        if ass == 1:
            ms = [2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 16, 16, 8]
            co = [1, 1, 1, 1, 2, 2, 2, 2, 3, 3, 1, 2, 3]
        elif ass == 2:
            ms = [2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 16, 16, 8]
            co = [4, 4, 4, 4, 5, 5, 5, 5, 3, 3, 4, 5, 3]
        elif ass == 3:
            if dep is None:
                raise ValueError("Department is required for MODEL exam.")
            if dep == 1:
                ms = [2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 16, 16, 16, 16, 16]
                co = [1, 1, 2, 2, 3, 3, 4, 4, 5, 5, 1, 2, 3, 4, 5]
            elif dep == 2:
                ms = [2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 13, 13, 13, 13, 13, 15]
                co = [1, 1, 2, 2, 3, 3, 4, 4, 5, 5, 1, 2, 3, 4, 5, 5]
            else:
                raise ValueError("Invalid department selected.")
        elif ass == 4:
            ms = [20, 20, 20, 20, 20]
            co = [1, 2, 3, 4, 5]
        elif ass == 5:
            ms = [20, 20, 20, 20, 20]
            co = [1, 2, 3, 4, 5]
        else:
            raise ValueError("Invalid assessment number for this regulation.")
    else:
        raise ValueError("Unsupported regulation. Use 13, 17, or 21.")

    if len(ms) != len(co):
        raise ValueError("Pattern error: ms and co length mismatch.")

    return ms, co


def _band_bounds(total: int, ms: List[int]) -> Tuple[List[int], List[int]]:
    """
    For a given total and question max marks, compute per-question [lo, hi] ranges
    using the original 'weak/average/good/excellent' logic.
    If impossible, it gracefully falls back to lo=0, hi=ms.
    """
    n = len(ms)
    lo = [0] * n
    hi = [0] * n

    if total == 0:
        return lo, hi

    for i, m in enumerate(ms):
        if total < 41:
            if total > 10:
                lo[i] = 0
                hi[i] = m
            elif total < 6:
                lo[i] = 0
                hi[i] = min(1, m)
            else:  # 6 <= total <= 10
                lo[i] = 0
                hi[i] = min(2, m)
        elif 40 < total < 61:
            if m < 3:
                lo[i] = 0
                hi[i] = m
            else:
                lo[i] = 5
                hi[i] = m
        elif 60 < total < 81:
            if m < 3:
                lo[i] = 1
                hi[i] = m
            else:
                lo[i] = 7
                hi[i] = m
        elif 80 < total < 100:
            if m < 3:
                lo[i] = 2
                hi[i] = m
            else:
                lo[i] = 9
                hi[i] = m
        else:
            lo[i] = m
            hi[i] = m

        lo[i] = max(0, min(lo[i], m))
        hi[i] = max(lo[i], min(hi[i], m))

    sum_lo = sum(lo)
    sum_hi = sum(hi)

    if not (sum_lo <= total <= sum_hi):
        lo = [0] * n
        hi = ms[:]

    return lo, hi


def random_split_total(total: int, ms: List[int]) -> List[int]:
    """
    Fast, guaranteed O(n) random split of 'total' into len(ms) parts,
    each within [0, ms[i]], using band-based min/max ranges.
    """
    if total < 0:
        raise ValueError("Total mark cannot be negative.")

    max_possible = sum(ms)
    if total > max_possible:
        raise ValueError(
            f"Total {total} exceeds maximum possible {max_possible} from split-up."
        )

    lo, hi = _band_bounds(total, ms)
    sum_lo = sum(lo)
    sum_hi = sum(hi)

    if not (sum_lo <= total <= sum_hi):
        raise RuntimeError("Internal bounds inconsistency.")

    n = len(ms)
    vals = lo[:]
    remaining = total - sum_lo
    caps = [hi[i] - lo[i] for i in range(n)]

    indices = list(range(n))
    ra.shuffle(indices)

    suffix_caps = [0] * (n + 1)
    for idx in range(n - 1, -1, -1):
        suffix_caps[idx] = suffix_caps[idx + 1] + caps[indices[idx]]

    for pos, i in enumerate(indices):
        if remaining <= 0:
            break

        max_for_this = caps[i]
        max_remaining_for_rest = suffix_caps[pos + 1]
        upper = min(max_for_this, remaining)
        min_extra = max(0, remaining - max_remaining_for_rest)

        if upper < min_extra:
            extra = min_extra
        else:
            extra = ra.randint(min_extra, upper)

        vals[i] += extra
        remaining -= extra

    if remaining != 0:
        raise RuntimeError("Random allocation failed to exhaust remaining marks.")

    return vals


def generate_assessment_csv(
    marks: List[int],
    reg: int,
    ass: int,
    dep: Optional[int],
    ass_name: str,
    custom_ms: Optional[List[int]] = None,
    custom_co: Optional[List[int]] = None,
) -> Tuple[bytes, str]:
    """
    Main engine for UI:
    - builds assessment pattern (or uses custom)
    - splits totals into question-wise marks
    - prepares CSV as bytes
    """
    if ass == 6:
        if not custom_ms or not custom_co:
            raise ValueError("Custom ms/co not provided for Custom assessment.")
        if len(custom_ms) != len(custom_co):
            raise ValueError("Custom ms and co length mismatch.")
        ms, co = custom_ms, custom_co
    else:
        ms, co = assessment_pattern(reg, ass, dep)

    qno = [i + 1 for i in range(len(ms))]

    all_splits: List[List[int]] = []
    for total in marks:
        split = random_split_total(total, ms)
        all_splits.append(split)

    # Prepare CSV in memory
    buffer = io.StringIO()
    writer = csv.writer(buffer)

    writer.writerow(["AI Spliiter By Sathish Ramanujam"])
    writer.writerow([f"Assessment Name : {ass_name}"])

    co1spup = [sum(ms[v] for v in range(len(ms)) if co[v] == 1)]
    co2spup = [sum(ms[v] for v in range(len(ms)) if co[v] == 2)]
    co3spup = [sum(ms[v] for v in range(len(ms)) if co[v] == 3)]
    co4spup = [sum(ms[v] for v in range(len(ms)) if co[v] == 4)]
    co5spup = [sum(ms[v] for v in range(len(ms)) if co[v] == 5)]

    writer.writerow(["QNO ->", *qno, "Course Outcome SPUP"])
    writer.writerow([
        "CO ->",
        *co,
        "co1Tot",
        "co2Tot",
        "co3Tot",
        "co4Tot",
        "co5Tot",
    ])
    writer.writerow([
        "TM | MS -> ",
        *ms,
        *co1spup,
        *co2spup,
        *co3spup,
        *co4spup,
        *co5spup,
    ])

    for spup in all_splits:
        co1tot = [spup[v] for v in range(len(spup)) if co[v] == 1]
        co2tot = [spup[v] for v in range(len(spup)) if co[v] == 2]
        co3tot = [spup[v] for v in range(len(spup)) if co[v] == 3]
        co4tot = [spup[v] for v in range(len(spup)) if co[v] == 4]
        co5tot = [spup[v] for v in range(len(spup)) if co[v] == 5]

        writer.writerow([
            sum(spup),
            *spup,
            sum(co1tot),
            sum(co2tot),
            sum(co3tot),
            sum(co4tot),
            sum(co5tot),
        ])

    csv_text = buffer.getvalue()
    buffer.close()
    filename = ass_name + ".csv"
    return csv_text.encode("utf-8"), filename


# ------------------ STREAMLIT UI ------------------ #

st.set_page_config(
    page_title="CO Split-Up Generator",
    layout="centered",
)

st.title("CO Split-Up Generator")
st.caption(" Designed by [Sathish Ramanujam](https://professor-sathish.github.io/) - Powered by IPS Tech Community ")

st.markdown("---")

# Upload
uploaded_file = st.file_uploader(
    "Upload marks CSV (first column = total marks)",
    type=["csv"],
    help="Typically your inputm.csv file.",
)

# Options
col1, col2 = st.columns(2)
with col1:
    reg = st.radio("Regulation", [13, 17, 21], index=2, horizontal=True)
with col2:
    ass_label_to_value = {
        "IA1": 1,
        "IA2": 2,
        "MODEL": 3,
        "Lab": 4,
        "Project": 5,
        "Custom": 6,
    }
    ass_label = st.selectbox(
        "Assessment Type",
        list(ass_label_to_value.keys()),
        index=0,
    )
    ass = ass_label_to_value[ass_label]

dep = None
custom_ms: Optional[List[int]] = None
custom_co: Optional[List[int]] = None

if ass == 3:
    dep_label = st.radio("Department (for MODEL exam)", ["S & H", "Other"], horizontal=True)
    dep = 1 if dep_label == "S & H" else 2

# Custom assessment configuration
if ass == 6:
    st.markdown("### Custom Assessment Configuration")
    q_count = st.number_input("Number of Questions", min_value=1, max_value=100, value=5, step=1)
    co_text = st.text_input(
        "CO number for each question (comma-separated)",
        value=",".join(str(i + 1) for i in range(min(q_count, 5))),
        help="Example for 5 questions: 1,2,3,4,5",
    )
    ms_text = st.text_input(
        "Max marks for each question (comma-separated)",
        value=",".join(["2"] * min(q_count, 5)),
        help="Example for 5 questions: 2,2,2,2,2",
    )

    try:
        co_list = [int(x.strip()) for x in co_text.split(",") if x.strip() != ""]
        ms_list = [int(x.strip()) for x in ms_text.split(",") if x.strip() != ""]
        if len(co_list) != q_count or len(ms_list) != q_count:
            st.warning(f"Please provide exactly {q_count} CO numbers and {q_count} max marks.")
        else:
            custom_co = co_list
            custom_ms = ms_list
    except ValueError:
        st.warning("CO numbers and max marks must be integers separated by commas.")

ass_name = st.text_input(
    "Output filename (without .csv)",
    value="assessment_output",
    max_chars=50,
)

st.markdown("---")

generate_btn = st.button("🚀 Generate CO Split-Up", type="primary")

if generate_btn:
    if uploaded_file is None:
        st.error("Please upload a CSV file with marks.")
    elif not ass_name.strip():
        st.error("Please enter a valid output filename.")
    elif ass == 6 and (custom_ms is None or custom_co is None):
        st.error("Please provide valid custom CO numbers and max marks.")
    else:
        try:
            marks = get_marks_from_file(uploaded_file)
            if not marks:
                st.error("No valid numeric marks found in the uploaded file.")
            else:
                csv_bytes, filename = generate_assessment_csv(
                    marks=marks,
                    reg=reg,
                    ass=ass,
                    dep=dep,
                    ass_name=ass_name.strip(),
                    custom_ms=custom_ms,
                    custom_co=custom_co,
                )

                st.success(
                    f"Generated CO split-up for {len(marks)} students as '{filename}'."
                )

                st.download_button(
                    label="⬇️ Download CSV",
                    data=csv_bytes,
                    file_name=filename,
                    mime="text/csv",
                )

                # Optional: quick peek of first few lines
                preview = csv_bytes.decode("utf-8").splitlines()[:15]
                st.markdown("**Preview (first 15 lines):**")
                st.code("\n".join(preview), language="text")

        except Exception as e:
            st.error(f"Error: {e}")
