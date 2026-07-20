from __future__ import annotations

import io
import os
import re
from datetime import datetime
from html.parser import HTMLParser
from pathlib import Path

import pandas as pd
from flask import Flask, jsonify, render_template, request, send_file
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter


app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 20 * 1024 * 1024

REQUIRED_COLUMNS = ["缴费日期", "业务类别", "预缴金额", "实缴金额"]
OUTPUT_COLUMNS = ["缴费日期按日汇总", "银行存款", "预收报名费", "预收报名费转收入", "收入", "税"]


class InputError(ValueError):
    pass


class ExcelHtmlTableParser(HTMLParser):
    """Small, dependency-free parser for WPS/Excel HTML exports."""

    def __init__(self):
        super().__init__(convert_charrefs=True)
        self.tables = []
        self._table = None
        self._row = None
        self._cell = None
        self._cell_colspan = 1

    def handle_starttag(self, tag, attrs):
        tag = tag.lower()
        if tag == "table" and self._table is None:
            self._table = []
        elif tag == "tr" and self._table is not None:
            self._row = []
        elif tag in ("td", "th") and self._row is not None:
            self._cell = []
            attrs = dict(attrs)
            try:
                self._cell_colspan = max(1, int(attrs.get("colspan", "1")))
            except ValueError:
                self._cell_colspan = 1
        elif tag == "br" and self._cell is not None:
            self._cell.append("\n")

    def handle_data(self, data):
        if self._cell is not None:
            self._cell.append(data)

    def handle_endtag(self, tag):
        tag = tag.lower()
        if tag in ("td", "th") and self._cell is not None and self._row is not None:
            value = "".join(self._cell).strip()
            self._row.append(value)
            self._row.extend([""] * (self._cell_colspan - 1))
            self._cell = None
            self._cell_colspan = 1
        elif tag == "tr" and self._row is not None and self._table is not None:
            if self._row:
                self._table.append(self._row)
            self._row = None
        elif tag == "table" and self._table is not None:
            if self._table:
                self.tables.append(self._table)
            self._table = None


def _read_excel_html(content: bytes) -> pd.DataFrame:
    text = content.decode("utf-8-sig", errors="replace")
    parser = ExcelHtmlTableParser()
    parser.feed(text)
    if not parser.tables:
        raise InputError("没有在文件中找到数据表。")
    table = max(parser.tables, key=lambda rows: len(rows) * max((len(r) for r in rows), default=0))
    width = max(len(row) for row in table)
    normalized = [row + [""] * (width - len(row)) for row in table]
    return pd.DataFrame(normalized)


def _clean_header(value) -> str:
    if pd.isna(value):
        return ""
    return re.sub(r"\s+", "", str(value)).strip()


def _normalize_dataframe(raw: pd.DataFrame) -> pd.DataFrame:
    raw = raw.dropna(how="all").dropna(axis=1, how="all")
    if raw.empty:
        raise InputError("表格没有可读取的数据。")

    # Find the row containing all required headers, allowing title rows above it.
    header_row = None
    for idx in range(min(len(raw), 30)):
        values = {_clean_header(v) for v in raw.iloc[idx].tolist()}
        if set(REQUIRED_COLUMNS).issubset(values):
            header_row = idx
            break

    if header_row is not None:
        headers = [_clean_header(v) for v in raw.iloc[header_row].tolist()]
        df = raw.iloc[header_row + 1 :].copy()
        df.columns = headers
    else:
        current = [_clean_header(v) for v in raw.columns]
        raw.columns = current
        df = raw.copy()

    missing = [c for c in REQUIRED_COLUMNS if c not in df.columns]
    if missing:
        raise InputError("缺少必要字段：" + "、".join(missing))

    # Duplicate header names can occur in exported spreadsheets; keep the first.
    df = df.loc[:, ~df.columns.duplicated()].copy()
    df["缴费日期"] = pd.to_datetime(df["缴费日期"], errors="coerce")
    df["业务类别"] = df["业务类别"].fillna("").astype(str).str.strip()
    for col in ["预缴金额", "实缴金额"]:
        cleaned = df[col].astype(str).str.replace(",", "", regex=False).str.replace("￥", "", regex=False).str.strip()
        df[col] = pd.to_numeric(cleaned, errors="coerce").fillna(0)

    df = df[df["缴费日期"].notna()].copy()
    if df.empty:
        raise InputError("“缴费日期”列没有有效日期。")
    df["日期"] = df["缴费日期"].dt.normalize()
    return df


def read_uploaded_excel(file_storage) -> pd.DataFrame:
    filename = (file_storage.filename or "").lower()
    content = file_storage.read()
    if not content:
        raise InputError("上传的文件为空。")

    stream = io.BytesIO(content)
    try:
        # Some WPS/ERP exports use an HTML table with an .xls extension.
        prefix = content[:500].lstrip().lower()
        if b"<html" in prefix or b"<table" in prefix:
            raw = _read_excel_html(content)
        elif filename.endswith(".xlsx") or filename.endswith(".xlsm"):
            raw = pd.read_excel(stream, sheet_name=0, header=None, engine="openpyxl")
        elif filename.endswith(".xls"):
            raw = pd.read_excel(stream, sheet_name=0, header=None, engine="xlrd")
        else:
            raise InputError("仅支持 .xls、.xlsx 或 .xlsm 文件。")
    except InputError:
        raise
    except Exception as exc:
        raise InputError(f"无法读取该 Excel：{exc}") from exc
    return _normalize_dataframe(raw)


def calculate_summary(df: pd.DataFrame) -> pd.DataFrame:
    daily = df.groupby("日期", as_index=False).agg(
        银行存款=("实缴金额", "sum"),
        预收报名费转收入=("预缴金额", "sum"),
    )
    prepaid = (
        df[df["业务类别"] == "预缴费"]
        .groupby("日期", as_index=False)["实缴金额"]
        .sum()
        .rename(columns={"实缴金额": "预收报名费"})
    )
    result = daily.merge(prepaid, on="日期", how="left")
    result["预收报名费"] = result["预收报名费"].fillna(0)
    result["收入"] = (result["银行存款"] - result["预收报名费"] + result["预收报名费转收入"]) / 1.03
    result["税"] = result["收入"] * 0.03
    result = result.rename(columns={"日期": "缴费日期按日汇总"})
    return result[OUTPUT_COLUMNS].sort_values("缴费日期按日汇总")


def make_workbook(summary: pd.DataFrame, source: pd.DataFrame) -> io.BytesIO:
    wb = Workbook()
    ws = wb.active
    ws.title = "缴费日期按日汇总"
    detail = wb.create_sheet("缴费明细")

    header_fill = PatternFill("solid", fgColor="F2F2F2")
    total_fill = PatternFill("solid", fgColor="E2F0D9")
    thin = Side(style="thin", color="D9D9D9")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    ws.append(OUTPUT_COLUMNS)
    for cell in ws[1]:
        cell.fill = header_fill
        cell.border = border
        cell.alignment = Alignment(horizontal="left")

    for row in summary.itertuples(index=False, name=None):
        ws.append(list(row))
    total_row = ws.max_row + 1
    ws.cell(total_row, 1, "合计")
    for col in range(2, 7):
        letter = get_column_letter(col)
        ws.cell(total_row, col, f"=SUM({letter}2:{letter}{total_row - 1})")
    for cell in ws[total_row]:
        cell.fill = total_fill
        cell.font = Font(bold=True, color="375623")
        cell.border = border

    for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
        row[0].number_format = "yyyy-mm-dd"
        for cell in row[1:]:
            cell.number_format = "#,##0.00"
            cell.border = border
    widths = [22, 17, 18, 23, 17, 15]
    for idx, width in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(idx)].width = width
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:F{ws.max_row - 1}"

    detail_columns = [c for c in ["缴费日期", "业务类别", "预缴金额", "实缴金额"] if c in source.columns]
    detail.append(detail_columns)
    for cell in detail[1]:
        cell.fill = header_fill
        cell.font = Font(bold=True)
        cell.border = border
    for values in source[detail_columns].itertuples(index=False, name=None):
        detail.append(list(values))
    for row in detail.iter_rows(min_row=2):
        row[0].number_format = "yyyy-mm-dd hh:mm"
        for cell in row:
            cell.border = border
        for cell in row[2:]:
            cell.number_format = "#,##0.00"
    for i, width in enumerate([20, 16, 18, 18], 1):
        detail.column_dimensions[get_column_letter(i)].width = width
    detail.freeze_panes = "A2"
    detail.auto_filter.ref = detail.dimensions

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output


@app.get("/")
def index():
    return render_template("index.html")


@app.post("/api/process")
def process_file():
    if "file" not in request.files:
        return jsonify({"ok": False, "error": "请选择 Excel 文件。"}), 400
    try:
        source = read_uploaded_excel(request.files["file"])
        summary = calculate_summary(source)
        output = make_workbook(summary, source)
        stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        return send_file(
            output,
            as_attachment=True,
            download_name=f"缴费日期会计汇总_{stamp}.xlsx",
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
    except InputError as exc:
        return jsonify({"ok": False, "error": str(exc)}), 400
    except Exception as exc:
        app.logger.exception("Processing failed")
        return jsonify({"ok": False, "error": f"处理失败：{exc}"}), 500


@app.get("/health")
def health():
    return {"status": "ok"}


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", 5000)), debug=True)
