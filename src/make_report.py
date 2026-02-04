import os
from datetime import date, timedelta
import pandas as pd
from pyhive import hive
from xlsxwriter.workbook import Workbook
from xlsxwriter.worksheet import Worksheet
import io


def _define_formats(workbook: Workbook) -> dict:
    """엑셀 서식"""
    return {
        "row1_base": workbook.add_format({"bg_color": "#A9A9A9", "bottom": 5}),
        "header": workbook.add_format(
            {
                "bg_color": "#A9A9A9",
                "font_color": "black",
                "bold": True,
                "bottom": 5,
                "align": "center",
            }
        ),
        "title": workbook.add_format(
            {
                "bg_color": "#A9A9A9",
                "font_color": "black",
                "bold": True,
                "font_size": 12,
                "bottom": 5,
            }
        ),
        "light_grey_bg": workbook.add_format({"bg_color": "#F2F2F2"}),
        "summary_label": workbook.add_format({"bg_color": "#F2F2F2", "bold": True}),
        "summary_value": workbook.add_format(
            {"bg_color": "#F2F2F2", "bold": True, "num_format": "#,##0"}
        ),
        "breakdown": workbook.add_format({"bold": True, "underline": True}),
        "group_header": workbook.add_format(
            {
                "bold": True,
                "fg_color": "#FFFF00",
                "align": "center",
                "valign": "vcenter",
            }
        ),
        "group_rev_bold": workbook.add_format({"bold": True, "num_format": "#,##0"}),
        "number": workbook.add_format({"num_format": "#,##0"}),
        "label": workbook.add_format({"align": "left"}),
        "bold_label": workbook.add_format({"bold": True}),
    }


def _write_data_rows(
    worksheet: Worksheet,
    start_row: int,
    data_source: pd.Series,
    weeks: list,
    formats: dict,
):
    """Ads revenue, CPM, Impressions"""
    # 레이블 작성
    worksheet.write(start_row + 1, 2, "Ads revenue", formats["label"])
    worksheet.write(start_row + 2, 2, "CPM", formats["label"])
    worksheet.write(start_row + 3, 2, "Impressions (implied)", formats["label"])

    # 주차별 데이터 작성
    for col_num, week in enumerate(weeks):
        rev = data_source.get(("ads_revenue", week), 0)
        imp = data_source.get(("impressions", week), 0)
        cpm = (rev / imp * 1000) if imp else 0

        # 0보다 큰 값만 셀에 표기
        if rev > 0:
            worksheet.write(start_row + 1, 3 + col_num, rev, formats["group_rev_bold"])
        if cpm > 0:
            worksheet.write(start_row + 2, 3 + col_num, cpm, formats["number"])
        if imp > 0:
            worksheet.write(start_row + 3, 3 + col_num, imp, formats["number"])


def create_weekly_report(db_config):
    print("스크립트 실행 시작...")

    # --- 1. 데이터 로딩 ---
    try:
        conn = hive.Connection(**db_config)

        # 스키마 변경과 함께 바뀔 예정
        query = "SELECT * from jaeyoung.coupang_play_report"
        df = pd.read_sql(query, conn)
        conn.close()
        print(
            f"데이터를 성공적으로 불러왔습니다. 총 {len(df)}개의 행이 로드되었습니다."
        )
    except Exception as e:
        raise ConnectionError(f"데이터 로딩 중 에러가 발생했습니다: {e}")

    if df.empty:
        print("데이터가 비어있습니다.")

    # --- 2. 데이터 전처리 ---
    print("데이터 전처리 시작...")
    df["ads_revenue"] = df["cpm"] * df["impressions"] // 1000
    weeks = sorted(df["w"].unique())

    # 데이터 최신 여부 검증
    last_week_date = date.today() - timedelta(days=7)
    thisweek = f"W{last_week_date.isocalendar()[1]}"
    sheetweek = weeks[-1] if weeks else ""
    if sheetweek and thisweek != sheetweek:
        raise ValueError(
            f"데이터 업데이트 지연: 실제 주차는 {thisweek}이지만, 데이터의 최신 주차는 {sheetweek}입니다."
        )

    # 피벗 테이블 및 그룹 합계 생성
    start_week_map = df.groupby("camp_nm")["w"].min()
    pivot = df.pivot_table(
        index=["ads_group_nm", "camp_nm"],
        columns="w",
        values=["ads_revenue", "impressions"],
        aggfunc="sum",
        fill_value=0,
    )
    group_totals = pivot.groupby(level="ads_group_nm").sum()

    # 그룹별 최신 주차 활성 캠페인 집합 생성
    active_campaigns = {}
    for group_name in df["ads_group_nm"].unique():
        group_df = df[df["ads_group_nm"] == group_name]
        if not group_df.empty:
            latest_week = group_df["w"].max()
            active_df = group_df[
                (group_df["w"] == latest_week) & (group_df["ads_revenue"] > 0)
            ]
            active_campaigns[group_name] = set(active_df["camp_nm"].unique())
    print("데이터 전처리 완료.")

    # --- 3. 엑셀 파일 생성 (인메모리) ---
    print("엑셀 파일 생성 시작...")
    today = date.today().strftime("%Y%m%d")
    thisweek_val = weeks[-1] if weeks else "W_NA"
    file_name = f"Nasmedia data request_{thisweek_val}_nap dsp_{today}_nasmedia.xlsx"

    output_buffer = io.BytesIO()
    writer = pd.ExcelWriter(output_buffer, engine="xlsxwriter")
    workbook = writer.book
    worksheet = workbook.add_worksheet("Nasmedia Data Request")
    formats = _define_formats(workbook)

    # --- 4. 엑셀 레이아웃 설정 ---
    worksheet.set_row(0, None, formats["row1_base"])
    worksheet.set_row(1, None, workbook.add_format({"bg_color": "white"}))
    for i in range(2, 5):
        worksheet.set_row(i, None, formats["light_grey_bg"])
    worksheet.set_column("A:B", 1.1)
    worksheet.set_column("C:C", 30)
    worksheet.freeze_panes("D2")

    worksheet.write("A1", "Ads Revenue", formats["title"])
    worksheet.write_row("B1", ["", ""], formats["row1_base"])
    worksheet.write("B7", "Revenue Breakdown", formats["breakdown"])

    # 주차별 컬럼 설정 (최근 10주만 표시)
    weeks_to_hide = weeks[:-10] if len(weeks) > 10 else []
    for i, week in enumerate(weeks):
        col_idx = 3 + i
        worksheet.write(0, col_idx, week, formats["header"])
        col_options = {"hidden": True} if week in weeks_to_hide else {}
        worksheet.set_column(col_idx, col_idx, 16, None, col_options)

    # Note 컬럼 설정
    note_col_idx = 3 + len(weeks)
    worksheet.write(0, note_col_idx, "Note", formats["header"])
    worksheet.set_column(note_col_idx, note_col_idx, 30)
    worksheet.write(
        2, note_col_idx, "Week definition: Mon-Sun", formats["summary_label"]
    )
    worksheet.write(
        3,
        note_col_idx,
        "Daily is fine too, if they are able to provide.",
        formats["summary_label"],
    )
    worksheet.write(
        4,
        note_col_idx,
        "Starting period: From July '28 (or go forward basis)",
        formats["summary_label"],
    )

    # --- 5. 데이터 작성 ---
    row = 8
    group_order = ["LIVE_SPORTS", "LIVE_FAST", "VOD"]
    for group_name in group_order:
        if group_name not in group_totals.index:
            continue

        active_camps_in_group = active_campaigns.get(group_name, set())

        # 그룹 헤더 및 데이터 작성
        worksheet.merge_range(row, 1, row, 2, group_name, formats["group_header"])
        _write_data_rows(worksheet, row, group_totals.loc[group_name], weeks, formats)
        row += 5

        # 캠페인 정렬: 활성 캠페인 -> 시작 주차 오름차순
        campaigns_in_group = pivot.loc[group_name].index
        sorted_campaigns = sorted(
            campaigns_in_group,
            key=lambda camp: (
                0 if camp in active_camps_in_group else 1,
                start_week_map.get(camp, "W99"),
            ),
        )

        for camp_name in sorted_campaigns:
            is_active = camp_name in active_camps_in_group
            # 비활성 캠페인인 경우, 할당된 5개의 행을 모두 숨김
            if not is_active:
                for i in range(5):
                    worksheet.set_row(row + i, None, None, {"hidden": True})

            # 캠페인 이름 및 데이터 작성 (숨겨진 행에도 데이터는 동일하게 작성)
            worksheet.write(row, 2, camp_name)
            _write_data_rows(
                worksheet, row, pivot.loc[(group_name, camp_name)], weeks, formats
            )
            row += 5

    # --- 6. 최종 집계 작성 ---
    worksheet.write("B3", "Total Ads Revenue", formats["summary_label"])
    worksheet.write("B4", "Average CPM", formats["summary_label"])
    worksheet.write("B5", "Total Nasmedia Fees", formats["summary_label"])

    for i, week in enumerate(weeks):
        col_idx = 3 + i
        weekly_revenue = group_totals.loc[:, ("ads_revenue", week)].sum()
        weekly_impressions = group_totals.loc[:, ("impressions", week)].sum()
        weekly_cpm = (
            (weekly_revenue / weekly_impressions * 1000) if weekly_impressions else 0
        )

        if weekly_revenue > 0:
            worksheet.write(2, col_idx, weekly_revenue, formats["summary_value"])
        if weekly_cpm > 0:
            worksheet.write(3, col_idx, weekly_cpm, formats["summary_value"])

    writer.close()
    print(f"'{file_name}' 파일이 인메모리로 성공적으로 생성되었습니다.")
    return file_name, output_buffer.getvalue()
