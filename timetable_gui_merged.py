"""
시간표 자동 생성기 GUI 병합본

- timetable-generator.py의 자동 시간표 생성/백트래킹 로직을 기본 뼈대로 사용
- leenamyung.py의 Tkinter GUI 흐름을 이 파일 목적에 맞게 단순화해 적용

실행 방법:
: subjectTable.xlsx 이름의 엑셀 파일을 직접 선택합니다.
"""

import os
import random
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

import pandas as pd


SCIENCE_OPTIONS = {
    "1": ("일반 물리학1", "일반물리학및실험1"),
    "2": ("화학1", "일반화학및실험1"),
    "3": ("생물학1", "일반생물학1"),
    "4": ("일반수학1및연습", "일반수학1및연습"),
    "5": ("기타 교양과목", None),
}

CORE_DOMAINS = ["랜덤 배정", "인문사회", "자연과학", "휴먼테크", "글로벌외국어", "과목명 검색"]
DAYS = ["월", "화", "수", "목", "금", "토", "일"]


def clean(value):
    if pd.isna(value):
        return ""
    return str(value).strip()


def parse_credit(value):
    try:
        return int(str(value).split("/")[-1])
    except Exception:
        return 0


class Subject:
    def __init__(self, row):
        self.name = clean(row.get("교과목명", ""))
        self.section = clean(row.get("분반", ""))
        self.credit = parse_credit(row.get("시간/학점", "0/0"))
        self.time = clean(row.get("요일및교시(강의실)", ""))
        self.prof = clean(row.get("담당교수", ""))
        self.type = clean(row.get("1전공기준 이수구분", ""))


def parse_slots(time_text):
    slots = set()
    if clean(time_text) == "":
        return slots

    for part in str(time_text).split(","):
        part = part.strip().split("(")[0]
        if len(part) < 2:
            continue

        day = part[0]
        time_part = part[1:]

        try:
            if "~" in time_part:
                start, end = map(int, time_part.split("~"))
                for period in range(start, end + 1):
                    slots.add((day, period))
            else:
                slots.add((day, int(time_part)))
        except Exception:
            pass

    return slots


def row_to_info(row):
    return Subject(row)


def unique_names(rows):
    result = []
    for row in rows:
        name = clean(row.get("교과목명", ""))
        if name and name not in result:
            result.append(name)
    return result


def find_combos(pools, target_credit=17, max_results=5):
    results = []

    def backtrack(idx, chosen, used_slots, used_names):
        if len(results) >= max_results:
            return

        if idx == len(pools):
            total = sum(item.credit for item in chosen)
            if total == target_credit:
                results.append(chosen[:])
            return

        pool = pools[idx][:]
        random.shuffle(pool)

        for subject in pool:
            if subject.name in used_names:
                continue

            slots = parse_slots(subject.time)
            if used_slots & slots:
                continue

            chosen.append(subject)
            used_names.add(subject.name)
            backtrack(idx + 1, chosen, used_slots | slots, used_names)
            chosen.pop()
            used_names.remove(subject.name)

    backtrack(0, [], set(), set())
    return results


def get_diversity_core_name(subjects):
    for subject in subjects:
        if subject.type == "중핵교양필수" and subject.name != "I-DESIGN":
            return subject.name
    return ""


def diversify_by_core_subject(results, max_results):
    selected = []
    used_core_names = set()
    used_signatures = set()

    def signature(subjects):
        return tuple(sorted((s.name, s.section, s.time) for s in subjects))

    for result in results:
        core_name = get_diversity_core_name(result)
        result_signature = signature(result)
        if result_signature in used_signatures:
            continue
        if core_name and core_name not in used_core_names:
            selected.append(result)
            used_core_names.add(core_name)
            used_signatures.add(result_signature)
        if len(selected) >= max_results:
            return selected

    for result in results:
        result_signature = signature(result)
        if result_signature in used_signatures:
            continue
        selected.append(result)
        used_signatures.add(result_signature)
        if len(selected) >= max_results:
            return selected

    return selected


def build_grid(subjects):
    grid = {}
    used_periods = set()

    for subject in subjects:
        for day, period in parse_slots(subject.time):
            if day in DAYS:
                grid[(day, period)] = subject.name
                used_periods.add(period)

    if not used_periods:
        used_periods = set(range(1, 10))

    return grid, sorted(used_periods)


def load_excel_with_dialog():
    folder = os.path.dirname(os.path.abspath(__file__))
    excel_path = os.path.join(folder, "subjectTable.xlsx")

    if os.path.exists(excel_path):
        return excel_path

    root = tk.Tk()
    root.withdraw()
    excel_path = filedialog.askopenfilename(
        title="subjectTable 엑셀 파일 선택",
        filetypes=[("Excel Files", "*.xlsx *.xls"), ("All Files", "*.*")],
    )
    root.destroy()
    return excel_path


def load_subject_table(excel_path):
    df = pd.read_excel(excel_path, dtype=str)
    df.columns = df.columns.str.strip()
    return df.dropna(how="all").reset_index(drop=True)


def get_major_names(df):
    if "비고" not in df.columns:
        return []

    majors = set()
    for value in df["비고"].dropna():
        for major in str(value).split(","):
            major = major.strip()
            if major:
                majors.add(major)
    return sorted(majors)


def prepare_generation(df, settings):
    target_major = settings["target_major"]
    off_day = settings["off_day"]

    rows = []
    for _, row in df.iterrows():
        remark = clean(row.get("비고", ""))
        time_text = clean(row.get("요일및교시(강의실)", ""))

        if remark and target_major not in [x.strip() for x in remark.split(",")]:
            continue
        if off_day and off_day in time_text:
            continue

        rows.append(row.to_dict())

    filtered = pd.DataFrame(rows, columns=df.columns)
    if filtered.empty:
        raise ValueError("선택한 전공/공강 조건에 맞는 과목을 찾지 못했습니다.")

    name_col = filtered["교과목명"].fillna("").str.strip()
    type_col = filtered["1전공기준 이수구분"].fillna("").str.strip()
    dept_col = filtered["개설전공(과)"].fillna("").str.strip()

    people_rows = filtered[name_col == "인간학1"].to_dict("records")
    design_rows = filtered[name_col == "I-DESIGN"].to_dict("records")
    core_rows = filtered[
        (type_col == "중핵교양필수") & (name_col != "I-DESIGN")
    ].to_dict("records")
    major_rows = filtered[
        (type_col == "전공기초") & (dept_col == target_major)
    ].to_dict("records")
    liberal_rows = filtered[type_col == "자유선택교양"].to_dict("records")

    core_choice = settings["core_choice"]
    if core_choice in {"인문사회", "자연과학", "휴먼테크", "글로벌외국어"}:
        core_pool = [
            r for r in core_rows
            if clean(r.get("중핵 편성영역", "")) == core_choice
        ]
        if not core_pool:
            core_pool = core_rows
    elif core_choice == "과목명 검색" and settings["core_keyword"]:
        keyword = settings["core_keyword"]
        core_pool = [
            r for r in core_rows
            if keyword in clean(r.get("교과목명", ""))
        ]
        if not core_pool:
            core_pool = core_rows
    else:
        core_pool = core_rows

    major_names = unique_names(major_rows)
    picked_major = settings["picked_major"]
    if not major_names:
        raise ValueError(f"[{target_major}] 전공기초 과목을 찾지 못했습니다.")
    if picked_major not in major_names:
        raise ValueError("선택한 전공기초 과목이 현재 조건의 과목 목록에 없습니다.")

    picked_options = settings["picked_options"]
    if len(picked_options) != 2 or len(set(picked_options)) != 2:
        raise ValueError("추가 과목은 서로 다른 항목 2개를 선택해야 합니다.")

    pools = []
    if people_rows:
        pools.append([row_to_info(r) for r in people_rows])
    if design_rows:
        pools.append([row_to_info(r) for r in design_rows])
    if core_pool:
        pools.append([row_to_info(r) for r in core_pool])

    picked_major_rows = [
        r for r in major_rows
        if clean(r.get("교과목명", "")) == picked_major
    ]
    pools.append([row_to_info(r) for r in picked_major_rows])

    for option in picked_options:
        subject_name = SCIENCE_OPTIONS[option][1]

        if subject_name is None:
            extra_rows = liberal_rows
        else:
            extra_rows = filtered[name_col == subject_name].to_dict("records")

        pools.append([row_to_info(r) for r in extra_rows])

    empty_pool_count = sum(1 for pool in pools if not pool)
    if empty_pool_count:
        raise ValueError("조건에 맞는 필수 과목 후보가 비어 있어 시간표를 만들 수 없습니다.")

    max_results = settings["max_results"]
    search_limit = max(max_results * 20, 80)
    raw_results = find_combos(
        pools,
        target_credit=settings["target_credit"],
        max_results=search_limit,
    )
    return diversify_by_core_subject(raw_results, max_results)


# =========================
# GUI 영역
# =========================

class SettingsWindow:
    def __init__(self, root, df):
        self.root = root
        self.df = df
        self.settings = None
        self.major_names = get_major_names(df)

        root.title("시간표 자동 생성 설정")
        root.geometry("560x560")
        root.resizable(False, False)

        main = ttk.Frame(root, padding=18)
        main.pack(fill="both", expand=True)

        ttk.Label(main, text="전공").pack(anchor="w")
        self.major_var = tk.StringVar(value=self.major_names[0] if self.major_names else "")
        self.major_combo = ttk.Combobox(
            main,
            textvariable=self.major_var,
            values=self.major_names,
            width=46,
        )
        self.major_combo.pack(fill="x", pady=(3, 12))
        self.major_combo.bind("<<ComboboxSelected>>", self.refresh_major_subjects)

        ttk.Label(main, text="공강 희망 요일").pack(anchor="w")
        self.off_day_var = tk.StringVar(value="")
        off_day_frame = ttk.Frame(main)
        off_day_frame.pack(fill="x", pady=(3, 12))
        for label, value in [("없음", ""), ("월", "월"), ("화", "화"), ("수", "수"), ("목", "목"), ("금", "금")]:
            ttk.Radiobutton(off_day_frame, text=label, value=value, variable=self.off_day_var).pack(side="left", padx=(0, 12))

        ttk.Label(main, text="추가 중핵교양필수 선택 방식").pack(anchor="w")
        self.core_choice_var = tk.StringVar(value=CORE_DOMAINS[0])
        self.core_combo = ttk.Combobox(
            main,
            textvariable=self.core_choice_var,
            values=CORE_DOMAINS,
            state="readonly",
            width=30,
        )
        self.core_combo.pack(fill="x", pady=(3, 6))

        self.core_keyword_var = tk.StringVar()
        ttk.Entry(main, textvariable=self.core_keyword_var).pack(fill="x", pady=(0, 12))
        ttk.Label(main, text="과목명 검색을 선택한 경우에만 검색어가 사용됩니다.").pack(anchor="w")

        ttk.Separator(main).pack(fill="x", pady=12)

        ttk.Label(main, text="듣고 싶은 전공기초 과목 1개").pack(anchor="w")
        self.major_subject_var = tk.StringVar()
        self.major_subject_combo = ttk.Combobox(
            main,
            textvariable=self.major_subject_var,
            state="readonly",
            width=46,
        )
        self.major_subject_combo.pack(fill="x", pady=(3, 12))

        ttk.Label(main, text="추가 과목 2개 선택").pack(anchor="w")
        self.option_vars = {}
        option_frame = ttk.Frame(main)
        option_frame.pack(fill="x", pady=(3, 12))
        for key, (label, _) in SCIENCE_OPTIONS.items():
            var = tk.BooleanVar(value=False)
            self.option_vars[key] = var
            ttk.Checkbutton(option_frame, text=label, variable=var).pack(anchor="w")

        bottom = ttk.Frame(main)
        bottom.pack(fill="x", pady=(10, 0))

        ttk.Label(bottom, text="추천 개수").pack(side="left")
        self.max_results_var = tk.IntVar(value=5)
        ttk.Spinbox(bottom, from_=1, to=20, textvariable=self.max_results_var, width=6).pack(side="left", padx=(6, 0))

        ttk.Button(main, text="시간표 생성", command=self.confirm).pack(fill="x", pady=(18, 0))

        self.refresh_major_subjects()

    def get_current_major_rows(self):
        target_major = self.major_var.get().strip()
        rows = []
        for _, row in self.df.iterrows():
            remark = clean(row.get("비고", ""))
            if remark and target_major not in [x.strip() for x in remark.split(",")]:
                continue
            rows.append(row.to_dict())

        filtered = pd.DataFrame(rows, columns=self.df.columns)
        if filtered.empty:
            return []

        name_col = filtered["교과목명"].fillna("").str.strip()
        type_col = filtered["1전공기준 이수구분"].fillna("").str.strip()
        dept_col = filtered["개설전공(과)"].fillna("").str.strip()

        major_rows = filtered[
            (type_col == "전공기초") & (dept_col == target_major)
        ].to_dict("records")
        return unique_names(major_rows)

    def refresh_major_subjects(self, event=None):
        major_subjects = self.get_current_major_rows()
        self.major_subject_combo.configure(values=major_subjects)
        self.major_subject_var.set(major_subjects[0] if major_subjects else "")

    def confirm(self):
        picked_options = [key for key, var in self.option_vars.items() if var.get()]
        if not self.major_var.get().strip():
            messagebox.showerror("입력 오류", "전공을 입력하거나 선택해주세요.")
            return
        if not self.major_subject_var.get().strip():
            messagebox.showerror("입력 오류", "전공기초 과목을 선택해주세요.")
            return
        if len(picked_options) != 2:
            messagebox.showerror("입력 오류", "추가 과목은 정확히 2개를 선택해주세요.")
            return

        self.settings = {
            "target_major": self.major_var.get().strip(),
            "off_day": self.off_day_var.get().strip(),
            "core_choice": self.core_choice_var.get().strip(),
            "core_keyword": self.core_keyword_var.get().strip(),
            "picked_major": self.major_subject_var.get().strip(),
            "picked_options": picked_options,
            "target_credit": 17,
            "max_results": int(self.max_results_var.get()),
        }
        self.root.destroy()


class ResultWindow:
    def __init__(self, root, results, settings):
        self.root = root
        self.results = results
        self.settings = settings
        self.current_idx = 0
        self.cell_labels = {}

        root.title("추천 시간표 결과")
        root.geometry("1120x760")
        root.resizable(True, True)

        self.left = ttk.Frame(root, padding=10)
        self.left.grid(row=0, column=0, sticky="nsew")

        self.right = ttk.Frame(root, padding=10)
        self.right.grid(row=0, column=1, sticky="nsew")

        root.grid_columnconfigure(0, weight=7)
        root.grid_columnconfigure(1, weight=3)
        root.grid_rowconfigure(0, weight=1)

        toolbar = ttk.Frame(self.left)
        toolbar.pack(fill="x", pady=(0, 8))

        self.title_label = ttk.Label(toolbar, text="", font=("맑은 고딕", 12, "bold"))
        self.title_label.pack(side="left")

        ttk.Button(toolbar, text="이전", command=self.prev_result).pack(side="right", padx=(6, 0))
        ttk.Button(toolbar, text="다음", command=self.next_result).pack(side="right")

        self.table_frame = ttk.Frame(self.left)
        self.table_frame.pack(fill="both", expand=True)
        self.init_table()

        ttk.Label(self.right, text="과목 요약", font=("맑은 고딕", 12, "bold")).pack(anchor="w")
        self.summary = tk.Text(self.right, wrap="word", width=38, height=30)
        self.summary.pack(fill="both", expand=True, pady=(8, 0))

        self.render()

    def init_table(self):
        ttk.Label(self.table_frame, text="교시", anchor="center").grid(row=0, column=0, sticky="nsew")
        for col, day in enumerate(DAYS, 1):
            ttk.Label(self.table_frame, text=day, anchor="center").grid(row=0, column=col, sticky="nsew")

        for row, period in enumerate(range(1, 10), 1):
            ttk.Label(self.table_frame, text=f"{period}교시", anchor="center").grid(row=row, column=0, sticky="nsew")
            for col, day in enumerate(DAYS, 1):
                label = tk.Label(
                    self.table_frame,
                    text="",
                    bg="white",
                    relief="solid",
                    bd=1,
                    wraplength=110,
                    justify="center",
                    font=("맑은 고딕", 9),
                )
                label.grid(row=row, column=col, sticky="nsew")
                self.cell_labels[(day, period)] = label

        for col in range(8):
            self.table_frame.grid_columnconfigure(col, weight=1, minsize=80)
        for row in range(10):
            self.table_frame.grid_rowconfigure(row, weight=1, minsize=54)

    def prev_result(self):
        if not self.results:
            return
        self.current_idx = (self.current_idx - 1) % len(self.results)
        self.render()

    def next_result(self):
        if not self.results:
            return
        self.current_idx = (self.current_idx + 1) % len(self.results)
        self.render()

    def render(self):
        for label in self.cell_labels.values():
            label.config(text="", bg="white")

        if not self.results:
            self.title_label.config(text="조건에 맞는 시간표 조합이 없습니다.")
            self.summary.delete("1.0", tk.END)
            self.summary.insert(tk.END, "공강 요일, 전공기초 과목, 추가 과목 조건을 바꿔 다시 시도해보세요.")
            return

        subjects = self.results[self.current_idx]
        total = sum(subject.credit for subject in subjects)
        self.title_label.config(
            text=f"추천 시간표 {self.current_idx + 1} / {len(self.results)} - 총 {total}학점"
        )

        colors = ["#ffced3", "#c9ead0", "#cfe4ff", "#ead8ff", "#ffe3bd", "#d9f0f0", "#f3efbf"]
        grid, _ = build_grid(subjects)
        subject_colors = {subject.name: colors[idx % len(colors)] for idx, subject in enumerate(subjects)}

        for (day, period), name in grid.items():
            if (day, period) in self.cell_labels:
                self.cell_labels[(day, period)].config(text=name, bg=subject_colors.get(name, "white"))

        self.summary.delete("1.0", tk.END)
        self.summary.insert(tk.END, f"전공: {self.settings['target_major']}\n")
        self.summary.insert(tk.END, f"공강 희망: {self.settings['off_day'] or '없음'}\n\n")
        for subject in subjects:
            line = (
                f"- [{subject.type}] {subject.name}\n"
                f"  분반: {subject.section} | {subject.credit}학점 | 교수: {subject.prof}\n"
                f"  시간: {subject.time}\n\n"
            )
            self.summary.insert(tk.END, line)


def main():
    excel_path = load_excel_with_dialog()
    if not excel_path:
        return

    try:
        df = load_subject_table(excel_path)
    except Exception as exc:
        messagebox.showerror("파일 오류", f"엑셀 파일을 읽을 수 없습니다.\n{exc}")
        return

    settings_root = tk.Tk()
    settings_window = SettingsWindow(settings_root, df)
    settings_root.mainloop()

    if not settings_window.settings:
        return

    try:
        results = prepare_generation(df, settings_window.settings)
    except Exception as exc:
        messagebox.showerror("생성 오류", str(exc))
        return

    result_root = tk.Tk()
    ResultWindow(result_root, results, settings_window.settings)
    result_root.mainloop()


if __name__ == "__main__":
    main()
