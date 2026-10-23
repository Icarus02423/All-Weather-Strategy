"""从同花顺 iFinD HTTP API 下载 12 个加权期货的日收盘价。

首次先运行 --probe 核验短区间；直接运行则下载 START_DATE 起的完整区间。
仅使用 Python 标准库，兼容 Python 3.9。导入模块不会读凭证、联网或写文件。
本地 .env 可保存账号、密码和令牌；当前 HTTP 流程仅读取 IFIND_REFRESH_TOKEN。
"""
from __future__ import annotations

import argparse
import csv
import json
import math
import os
import re
import socket
import ssl
import sys
import tempfile
import time
from dataclasses import dataclass
from datetime import date, datetime, timedelta, timezone
from http.client import HTTPException
from pathlib import Path
from typing import Any, Iterator, Mapping
from urllib.error import HTTPError, URLError
from urllib.request import HTTPSHandler, HTTPRedirectHandler, Request, build_opener

# 日常修改日期和品种时，优先修改这一段。
START_DATE = "2013-09-06"
END_DATE = "2026-09-30"  # 留空：上海时间的昨天；也可填写 YYYY-MM-DD。
PROJECT_ROOT = Path(__file__).resolve().parent.parent
OUTPUT_DIR = PROJECT_ROOT / "data" / "futures"
SHANGHAI_TZ = timezone(timedelta(hours=8))
API_BASE = "https://quantapi.51ifind.com/api/v1"
TOKEN_KEY = "IFIND_REFRESH_TOKEN"
HISTORY_PARAMETERS = {"Interval": "D", "Fill": "Omit"}
OUTPUT_COLUMNS = ("日期", "代码", "名称", "资产类别", "收盘价")
ASSET_DIRECTORIES = {
    "股票": "stocks",
    "国债": "bonds",
    "黄金": "commodities",
    "其他商品": "commodities",
}
MANIFEST_COLUMNS = (
    "代码", "名称", "资产类别", "上市时间", "设定开始日期", "设定结束日期",
    "请求开始日期", "请求结束日期", "实际开始日期", "实际结束日期", "数据行数",
    "状态", "对齐日期缺失数", "非正收盘价数", "数据来源", "下载时间",
    "序列类型", "数据频率", "缺失填充方式",
)
PROBE_COLUMNS = (
    "代码", "名称", "资产类别", "核验区间", "请求开始日期", "请求结束日期",
    "实际开始日期", "实际结束日期", "数据行数", "状态", "错误信息",
)
PROBE_SAMPLE_COLUMNS = ("日期", "代码", "名称", "资产类别", "核验区间", "收盘价")


@dataclass(frozen=True)
class Instrument:
    code: str
    name: str
    asset_class: str
    listed_on: str


# 原油属于能源中心，但 iFinD 的这条加权序列使用 SHF 后缀。
INSTRUMENTS = (
    Instrument("IF8888.CFE", "沪深300加权", "股票", "2010-04-16"),
    Instrument("IC8888.CFE", "中证500加权", "股票", "2015-04-16"),
    Instrument("IM8888.CFE", "中证1000加权", "股票", "2022-07-22"),
    Instrument("TS8888.CFE", "2年国债加权", "国债", "2018-08-17"),
    Instrument("TF8888.CFE", "5年国债加权", "国债", "2013-09-06"),
    Instrument("T8888.CFE", "10年国债加权", "国债", "2015-03-20"),
    Instrument("TL8888.CFE", "30年国债加权", "国债", "2023-04-21"),
    Instrument("AU8888.SHF", "沪金加权", "黄金", "2008-01-09"),
    Instrument("SC8888.SHF", "原油加权", "其他商品", "2018-03-26"),
    Instrument("CU8888.SHF", "沪铜加权", "其他商品", "1993-03"),
    Instrument("RB8888.SHF", "螺纹钢加权", "其他商品", "2009-03-27"),
    Instrument("M8888.DCE", "豆粕加权", "其他商品", "2000-07-17"),
)


def parse_date(value: str) -> date:
    """配置日期严格使用 YYYY-MM-DD。"""
    if not isinstance(value, str) or not re.fullmatch(r"\d{4}-\d{2}-\d{2}", value):
        raise ValueError("日期必须使用 YYYY-MM-DD 格式。")
    return date.fromisoformat(value)

def load_refresh_token(env_file: Path) -> str:
    """只解析所需变量，不执行 .env，也不打印其中的内容。"""
    token = os.environ.get(TOKEN_KEY, "").strip()
    if not token and env_file.is_file():
        for line in env_file.read_text(encoding="utf-8-sig").splitlines():
            line = line.strip()
            if line.startswith("export "):
                line = line[7:].lstrip()
            key, separator, value = line.partition("=")
            if not separator or key.strip() != TOKEN_KEY:
                continue
            value = value.strip()
            if value.startswith(("'", '"')):
                match = re.fullmatch(r"(['\"])(.*?)\1(?:\s*#.*)?", value)
                if match is None:
                    raise ValueError(".env 中 IFIND_REFRESH_TOKEN 的引号不完整。")
                token = match.group(2).strip()
            else:
                token = re.split(r"\s+#", value, maxsplit=1)[0].strip()
    if not token:
        raise ValueError("未找到 IFIND_REFRESH_TOKEN，请在项目根目录 .env 中配置。")
    if "\r" in token or "\n" in token:
        raise ValueError("IFIND_REFRESH_TOKEN 不能包含换行。")
    return token

class _NoRedirect(HTTPRedirectHandler):
    def redirect_request(self, req, fp, code, msg, headers, newurl):
        # 令牌仅发送到固定的同花顺接口，不随重定向转发。
        return None

class IFindClient:
    """标准库 HTTP 客户端；仅对网络错误和 HTTP 429/5xx 有限重试。"""

    def __init__(self, timeout: float = 60, attempts: int = 3, retry_delay: float = 3):
        if timeout <= 0 or attempts < 1 or retry_delay < 0:
            raise ValueError("超时和尝试次数必须为正，重试间隔不能为负。")
        self.timeout = timeout
        self.attempts = attempts
        self.retry_delay = retry_delay
        context = ssl.create_default_context()
        # macOS 上沿用系统证书库，始终验证服务端证书。
        for cafile in (Path("/etc/ssl/cert.pem"), Path("/private/etc/ssl/cert.pem")):
            if cafile.is_file():
                context.load_verify_locations(cafile=str(cafile))
                break
        self.opener = build_opener(HTTPSHandler(context=context), _NoRedirect())
        self.access_token = ""
        self._secrets: list[str] = []

    def _post(self, endpoint: str, body: Mapping[str, Any], headers: Mapping[str, str]) -> dict:
        request = Request(
            f"{API_BASE}/{endpoint}",
            data=json.dumps(body).encode("utf-8"),
            headers={"Content-Type": "application/json", "ifindlang": "cn", **headers},
            method="POST",
        )
        for attempt in range(1, self.attempts + 1):
            try:
                with self.opener.open(request, timeout=self.timeout) as response:
                    raw = response.read()
            except HTTPError as exc:
                retryable = exc.code == 429 or 500 <= exc.code < 600
                if exc.fp is not None:
                    exc.close()
                if not retryable or attempt == self.attempts:
                    raise RuntimeError(f"同花顺请求失败（HTTP {exc.code}），请检查接口权限及网络。") from None
            except (URLError, TimeoutError, socket.timeout, ConnectionError, HTTPException) as exc:
                reason = getattr(exc, "reason", exc)
                if isinstance(reason, ssl.SSLCertVerificationError):
                    raise RuntimeError("HTTPS证书验证失败，请检查Python或系统CA证书配置。") from None
                if attempt == self.attempts:
                    raise RuntimeError("同花顺请求多次网络失败，请稍后重试。") from None
            else:
                try:
                    payload = json.loads(raw.decode("utf-8-sig"))
                except (UnicodeDecodeError, json.JSONDecodeError):
                    raise RuntimeError("同花顺响应不是有效的JSON，已有数据文件未修改。") from None
                if not isinstance(payload, dict) or "errorcode" not in payload:
                    raise RuntimeError("同花顺响应缺少有效的 errorcode。")
                if str(payload["errorcode"]) != "0":
                    detail = str(payload.get("errmsg", "请检查参数、令牌和数据权限。"))
                    for secret in self._secrets:
                        detail = detail.replace(secret, "[REDACTED]")
                    detail = " ".join(detail.split())[:200]
                    errorcode = str(payload["errorcode"])
                    if not re.fullmatch(r"-?\d+", errorcode):
                        errorcode = "unknown"
                    raise RuntimeError(f"同花顺业务错误 {errorcode}：{detail}")
                return payload
            print(
                f"⚠️ 同花顺第 {attempt}/{self.attempts} 次请求暂时失败，"
                f"{self.retry_delay * attempt:g} 秒后重试...",
                flush=True,
            )
            time.sleep(self.retry_delay * attempt)
        raise RuntimeError("同花顺请求未完成。")

    def login(self, refresh_token: str) -> None:
        if not refresh_token or "\r" in refresh_token or "\n" in refresh_token:
            raise ValueError("IFIND_REFRESH_TOKEN 不能为空或包含换行。")
        self._secrets.append(refresh_token)
        payload = self._post("get_access_token", {}, {"refresh_token": refresh_token})
        data = payload.get("data")
        access_token = data.get("access_token") if isinstance(data, dict) else None
        if not isinstance(access_token, str) or not access_token.strip() or "\r" in access_token or "\n" in access_token:
            raise RuntimeError("同花顺鉴权响应缺少有效的 access_token。")
        self.access_token = access_token.strip()
        self._secrets.append(self.access_token)

    def post(self, endpoint: str, body: Mapping[str, Any]) -> dict:
        if not self.access_token:
            raise RuntimeError("请先完成同花顺鉴权。")
        return self._post(endpoint, body, {"access_token": self.access_token})

def _table_field(table: Mapping[str, Any], field: str) -> list:
    for key, value in table.items():
        if str(key).lower() == field.lower():
            if isinstance(value, list):
                return value
            break
    raise RuntimeError(f"同花顺响应字段 {field} 缺失或不是列表。")

def split_date_range(start: date, end: date, days: int = 365) -> Iterator[tuple[date, date]]:
    if days < 1 or start > end:
        raise ValueError("分段天数必须为正，且开始日期不能晚于结束日期。")
    current = start
    while current <= end:
        last = min(current + timedelta(days=days - 1), end)
        yield current, last
        current = last + timedelta(days=1)

def _market_date(value: Any) -> date:
    text = str(value).strip()
    for fmt in ("%Y-%m-%d", "%Y%m%d", "%Y/%m/%d", "%Y-%m-%d %H:%M:%S"):
        try:
            return datetime.strptime(text, fmt).date()
        except ValueError:
            pass
    raise ValueError("同花顺行情中包含无法识别的日期。")



def effective_range(instrument: Instrument, start: date, end: date) -> tuple[date, date]:
    """上市时间仅精确到月份时，以月初作为请求裁剪下界。"""
    listing = instrument.listed_on
    if re.fullmatch(r"\d{4}-\d{2}", listing):
        listing += "-01"
    return max(start, parse_date(listing)), end


def parse_close_rows(payload: Mapping[str, Any], code: str, start: date, end: date) -> dict[date, float]:
    tables = payload.get("tables")
    if not isinstance(tables, list):
        raise RuntimeError(f"{code} 的行情响应缺少有效 tables。")
    rows: dict[date, float] = {}
    for item in tables:
        if not isinstance(item, dict) or str(item.get("thscode", "")).strip().upper() != code:
            raise RuntimeError(f"{code} 的响应代码不一致或表结构无效。")
        times, table = item.get("time"), item.get("table")
        if not isinstance(times, list) or not isinstance(table, dict):
            raise RuntimeError(f"{code} 的响应缺少日期列表或收盘价表。")
        prices = _table_field(table, "close")
        if len(times) != len(prices):
            raise RuntimeError(f"{code} 的日期和收盘价数量不一致。")
        for raw_day, raw_price in zip(times, prices):
            day = _market_date(raw_day)
            if not start <= day <= end:
                raise ValueError(f"{code} 返回请求区间外的日期 {day}。")
            try:
                if isinstance(raw_price, bool):
                    raise ValueError
                value = float(str(raw_price).strip().replace(",", ""))
                if not math.isfinite(value):
                    raise ValueError
            except (TypeError, ValueError):
                raise ValueError(f"{code} 在 {day} 的收盘价缺失或无效；不会补成0。") from None
            if day in rows and rows[day] != value:
                raise ValueError(f"{code} 在 {day} 出现冲突的重复收盘价。")
            rows[day] = value
    return rows


def fetch_close(client: IFindClient, instrument: Instrument, first: date, last: date) -> dict[date, float]:
    payload = client.post("cmd_history_quotation", {
        "codes": instrument.code,
        "indicators": "close",
        "startdate": first.isoformat(),
        "enddate": last.isoformat(),
        "functionpara": dict(HISTORY_PARAMETERS),
    })
    return parse_close_rows(payload, instrument.code, first, last)


def check_coverage(code: str, first: date, last: date, rows: Mapping[date, float]) -> None:
    """较长区间须覆盖两端，防止把只返回近五年的数据误判为完整历史。

    20 个自然日是异常提示阈值，不替代交易日历，也不证明逐日完整。
    """
    if not rows:
        # 短区间可能全是周末或长假；整段长于 20 天却全空需要排查。
        if (last - first).days >= 20:
            raise ValueError(f"{code} 在上市后区间 {first} 至 {last} 整段无行情；请检查代码和历史权限。")
        return
    if (min(rows) - first).days > 20 or (last - max(rows)).days > 20:
        raise ValueError(f"{code} 请求 {first} 至 {last}，实际仅 {min(rows)} 至 {max(rows)}；历史覆盖异常，停止保存。")


def download_instrument(client: IFindClient, instrument: Instrument, first: date, last: date) -> dict[date, float]:
    result: dict[date, float] = {}
    segments = list(split_date_range(first, last))
    for position, (begin, finish) in enumerate(segments, 1):
        print(f"  正在下载收盘价：{begin} 至 {finish}（{position}/{len(segments)}）", flush=True)
        chunk = fetch_close(client, instrument, begin, finish)
        check_coverage(instrument.code, begin, finish, chunk)
        for day, price in chunk.items():
            if day in result and result[day] != price:
                raise ValueError(f"{instrument.code} 在 {day} 出现跨段重复冲突。")
            result[day] = price
        print(f"  ✅ 本段取得 {len(chunk)} 个交易日", flush=True)
    if not result:
        raise ValueError(f"{instrument.code} 未取得行情；请检查是否处于有效交易日期区间。")
    check_coverage(instrument.code, first, last, result)
    return result


def write_csv_atomic(path: Path, columns: tuple[str, ...], rows) -> None:
    """文件写好后再替换；写入失败保留原文件。"""
    path.parent.mkdir(parents=True, exist_ok=True)
    temporary = None
    try:
        with tempfile.NamedTemporaryFile("w", encoding="utf-8-sig", newline="", dir=path.parent,
                                         prefix=f".{path.name}.", suffix=".tmp", delete=False) as stream:
            temporary = Path(stream.name)
            writer = csv.writer(stream)
            writer.writerow(columns)
            writer.writerows(rows)
            stream.flush()
            os.fsync(stream.fileno())
        os.replace(temporary, path)
    finally:
        if temporary is not None:
            temporary.unlink(missing_ok=True)





def make_summary(instrument: Instrument, first: date, last: date, rows: Mapping[date, float]) -> dict:
    return {
        "code": instrument.code, "name": instrument.name,
        "asset_class": instrument.asset_class, "listed_on": instrument.listed_on,
        "requested_start": first.isoformat(), "requested_end": last.isoformat(),
        "rows": len(rows), "actual_start": min(rows).isoformat() if rows else "",
        "actual_end": max(rows).isoformat() if rows else "",
        "nonpositive_price_dates": [day.isoformat() for day in sorted(rows) if rows[day] <= 0],
    }


def save_download(output: Path, results: list[tuple[Instrument, dict[date, float]]], manifest: dict) -> None:
    """整批下载检查成功后，保存收盘价总表、分类行情和下载清单。"""
    for instrument, _ in results:
        if instrument.asset_class not in ASSET_DIRECTORIES:
            raise ValueError(f"未配置资产类别的输出目录：{instrument.asset_class}。")
    union = sorted({day for _, rows in results for day in rows})
    overview = []
    for entry, (instrument, rows) in zip(manifest["instruments"], results):
        first, last = date.fromisoformat(entry["requested_start"]), date.fromisoformat(entry["requested_end"])
        missing = [day for day in union if first <= day <= last and day not in rows]
        not_listed = entry["status"] == "not_yet_listed"
        overview.append((
            instrument.code, instrument.name, instrument.asset_class, instrument.listed_on,
            manifest["requested_start"], manifest["requested_end"],
            "" if not_listed else entry["requested_start"], "" if not_listed else entry["requested_end"],
            entry["actual_start"], entry["actual_end"], len(rows), "尚未上市" if not_listed else "已下载",
            len(missing), len(entry["nonpositive_price_dates"]), manifest["source"], manifest["downloaded_at"],
            "加权期货价格", manifest["request_parameters"]["Interval"], manifest["request_parameters"]["Fill"],
        ))
    for instrument, rows in results:
        write_csv_atomic(output / ASSET_DIRECTORIES[instrument.asset_class] / f"{instrument.code}.csv", OUTPUT_COLUMNS, (
            (day.isoformat(), instrument.code, instrument.name, instrument.asset_class, format(rows[day], ".15g"))
            for day in sorted(rows)
        ))
    write_csv_atomic(output / "close_price.csv", ("日期", *(item.code for item, _ in results)), (
        (day.isoformat(), *(format(rows[day], ".15g") if day in rows else "" for _, rows in results))
        for day in union
    ))
    write_csv_atomic(output / "download_manifest.csv", MANIFEST_COLUMNS, overview)


def run_probe(client: IFindClient, instruments: tuple[Instrument, ...], start: date, end: date, report_path: Path) -> bool:
    """核验早期和近期窗口，分别输出扁平 CSV 报告及全部小样本行情。"""
    report_rows, sample_rows = [], []
    succeeded = True
    for instrument in instruments:
        first, last = effective_range(instrument, start, end)
        if first > last:
            report_rows.append((instrument.code, instrument.name, instrument.asset_class, "设定区间",
                                "", "", "", "", 0, "尚未上市", ""))
            continue
        windows = list(dict.fromkeys([(first, min(first + timedelta(days=13), last)),
                                      (max(first, last - timedelta(days=13)), last)]))
        for position, (begin, finish) in enumerate(windows):
            label = "最早/最近区间" if len(windows) == 1 else ("最早区间" if position == 0 else "最近区间")
            print(f"核验 {instrument.name}（{instrument.code}）：{label}，{begin} 至 {finish}", flush=True)
            actual_start, actual_end, count, status, error = "", "", 0, "失败", ""
            try:
                rows = fetch_close(client, instrument, begin, finish)
                if not rows:
                    raise ValueError("短区间无行情，请核对交易日、代码和历史权限。")
                actual_start, actual_end, count, status = min(rows).isoformat(), max(rows).isoformat(), len(rows), "可查询"
                sample_rows.extend((day.isoformat(), instrument.code, instrument.name, instrument.asset_class, label,
                                    format(rows[day], ".15g")) for day in sorted(rows))
                print(f"✅ 可查询，{count} 个交易日，{actual_start} 至 {actual_end}", flush=True)
            except (ValueError, RuntimeError) as exc:
                error, succeeded = str(exc), False
                print(f"❌ {instrument.code} 核验失败：{error}", flush=True)
            report_rows.append((instrument.code, instrument.name, instrument.asset_class, label,
                                begin.isoformat(), finish.isoformat(), actual_start, actual_end, count, status, error))
    write_csv_atomic(report_path, PROBE_COLUMNS, report_rows)
    sample_path = report_path.with_name("probe_samples.csv")
    write_csv_atomic(sample_path, PROBE_SAMPLE_COLUMNS, sample_rows)
    print(f"核验报告：{report_path}\n小样本行情：{sample_path}", flush=True)
    return succeeded


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="同花顺 iFinD 加权期货日收盘价下载工具。")
    parser.add_argument("--start", default=START_DATE, help="开始日期 YYYY-MM-DD。")
    parser.add_argument("--end", default=END_DATE, help="结束日期；留空默认上海时间昨天。")
    parser.add_argument("--codes", nargs="+", help="仅处理指定配置代码；默认全部12个。")
    parser.add_argument("--output", type=Path, default=OUTPUT_DIR, help="本地输出目录。")
    parser.add_argument("--list", action="store_true", help="离线列出代码，不读取凭证、不联网。")
    parser.add_argument("--probe", action="store_true", help="核验最早和最近的小样本，不生成正式行情CSV。")
    args = parser.parse_args(argv)
    started = time.perf_counter()
    try:
        start = parse_date(args.start)
        today = datetime.now(SHANGHAI_TZ).date()
        end = parse_date(args.end) if args.end else today - timedelta(days=1)
        if start > end or end >= today:
            raise ValueError("开始日期应不晚于结束日期；结束日期须早于上海时间今天。")
        codes = {code.strip().upper() for code in args.codes} if args.codes else {item.code for item in INSTRUMENTS}
        unknown = codes - {item.code for item in INSTRUMENTS}
        if unknown:
            raise ValueError(f"代码未在品种配置中：{', '.join(sorted(unknown))}。")
        selected = tuple(item for item in INSTRUMENTS if item.code in codes)
        output = args.output.expanduser().resolve()
        if args.codes and codes != {item.code for item in INSTRUMENTS} and output == OUTPUT_DIR.resolve():
            output = OUTPUT_DIR.parent / "futures_subset"
        print("同花顺 iFinD 加权期货数据下载工具\n" + "=" * 60, flush=True)
        print(f"数据来源：同花顺 iFinD HTTP API\n请求区间：{start} 至 {end}", flush=True)
        print("字段：日收盘价；不填补缺失值；接口交易日期保留原样。", flush=True)
        for item in selected:
            print(f"  {item.asset_class}：{item.name}（{item.code}）", flush=True)
        if args.list:
            return 0
        token = load_refresh_token(PROJECT_ROOT / ".env")
        client = IFindClient()
        print("正在登录同花顺...", flush=True)
        client.login(token)
        print("✅ 同花顺鉴权成功", flush=True)
        if args.probe:
            return 0 if run_probe(client, selected, start, end, output.parent / f"{output.name}_probe" / "probe_report.csv") else 1
        results = []
        manifest = {"source": "同花顺iFinD HTTP API", "series_type": "weighted_futures_price",
                    "request_parameters": dict(HISTORY_PARAMETERS), "requested_start": start.isoformat(),
                    "requested_end": end.isoformat(), "downloaded_at": datetime.now(SHANGHAI_TZ).isoformat(),
                    "instruments": [], "note": "加权价格序列，不自动等于可交易期货组合的收益序列。"}
        for position, instrument in enumerate(selected, 1):
            first, last = effective_range(instrument, start, end)
            print(f"\n[{position}/{len(selected)}] {instrument.name}（{instrument.code}）", flush=True)
            rows = {} if first > last else download_instrument(client, instrument, first, last)
            entry = make_summary(instrument, first, last, rows)
            entry["status"] = "not_yet_listed" if first > last else "downloaded"
            manifest["instruments"].append(entry)
            results.append((instrument, rows))
            if entry["nonpositive_price_dates"]:
                print("⚠️ 包含非正收盘价，已保留；计算百分比收益率前需另行处理。", flush=True)
        save_download(output, results, manifest)
        print("\n" + "=" * 60 + f"\n✅ 本次行情下载及保存完成：{output}", flush=True)
        for instrument, rows in results:
            actual = f"{min(rows)} 至 {max(rows)}" if rows else "请求区间尚未上市"
            print(f"{instrument.code}：{len(rows)} 行，{actual}", flush=True)
        print(f"总耗时：{time.perf_counter() - started:.1f} 秒", flush=True)
        return 0
    except (ValueError, RuntimeError, OSError) as exc:
        print(f"❌ 同花顺期货数据处理失败：{exc}", file=sys.stderr, flush=True)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
